import { useEffect, useMemo, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { countMatches, highlightText } from '../utils/highlight';
import { isSafeNavigationHref } from '../utils/sanitize';

interface RtfInlineStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  color?: string;
  backgroundColor?: string;
  verticalAlign?: 'sub' | 'super';
}

interface RtfParserState {
  destination?: string;
  linkHref?: string;
  skip: boolean;
  style: RtfInlineStyle;
  ucSkip: number;
}

interface RtfTextRun {
  type: 'text';
  text: string;
  style: RtfInlineStyle;
  href?: string;
}

interface RtfBreakRun {
  type: 'break';
}

type RtfInline = RtfTextRun | RtfBreakRun;

interface RtfParagraphBlock {
  type: 'paragraph';
  children: RtfInline[];
}

interface RtfTableCell {
  children: RtfInline[];
}

interface RtfTableBlock {
  type: 'table';
  rows: RtfTableCell[][];
}

type RtfBlock = RtfParagraphBlock | RtfTableBlock;

interface RtfPageMetrics {
  heightTwips?: number;
  marginBottomTwips?: number;
  marginLeftTwips?: number;
  marginRightTwips?: number;
  marginTopTwips?: number;
  widthTwips?: number;
}

interface ParsedRtfDocument {
  blocks: RtfBlock[];
  pageStyle?: CSSProperties;
  plainText: string;
}

const twipsPerInch = 1440;
const skippedDestinations = new Set([
  'colortbl',
  'datastore',
  'filetbl',
  'fonttbl',
  'generator',
  'info',
  'listoverridetable',
  'listtable',
  'object',
  'pict',
  'revtbl',
  'stylesheet',
  'themedata',
  'xmlnstbl',
]);

function decodeRtfBuffer(arrayBuffer: ArrayBuffer): string {
  try {
    return new TextDecoder('windows-1252').decode(arrayBuffer);
  } catch {
    return new TextDecoder('utf-8').decode(arrayBuffer);
  }
}

function decodeRtfByte(byte: number): string {
  try {
    return new TextDecoder('windows-1252').decode(new Uint8Array([byte]));
  } catch {
    return String.fromCharCode(byte);
  }
}

function cloneStyle(style: RtfInlineStyle): RtfInlineStyle {
  return { ...style };
}

function cloneState(state: RtfParserState): RtfParserState {
  return {
    ...state,
    style: cloneStyle(state.style),
  };
}

function styleKey(style: RtfInlineStyle, href?: string): string {
  return JSON.stringify({ ...style, href });
}

function extractColorTable(rtf: string): Map<number, string> {
  const colors = new Map<number, string>();
  const match = /{\\colortbl([\s\S]*?)}/.exec(rtf);

  if (!match) {
    return colors;
  }

  match[1].split(';').forEach((entry, index) => {
    const red = /\\red(\d+)/.exec(entry)?.[1];
    const green = /\\green(\d+)/.exec(entry)?.[1];
    const blue = /\\blue(\d+)/.exec(entry)?.[1];

    if (red && green && blue) {
      colors.set(index, `rgb(${Number(red)}, ${Number(green)}, ${Number(blue)})`);
    }
  });

  return colors;
}

function normalizeUnicodeValue(value: number): string {
  const codePoint = value < 0 ? value + 65536 : value;

  try {
    return String.fromCodePoint(codePoint);
  } catch {
    return '';
  }
}

function hrefFromFieldInstruction(instruction: string): string | undefined {
  const match = /\bHYPERLINK\s+(?:"([^"]+)"|([^\s]+))/i.exec(instruction);
  const href = match?.[1] ?? match?.[2];

  return href && isSafeNavigationHref(href) ? href : undefined;
}

function plainTextFromBlocks(blocks: RtfBlock[]): string {
  return blocks
    .map((block) => {
      if (block.type === 'table') {
        return block.rows
          .map((row) => row.map((cell) => cell.children.map((child) => (child.type === 'text' ? child.text : '\n')).join('')).join('\t'))
          .join('\n');
      }

      return block.children.map((child) => (child.type === 'text' ? child.text : '\n')).join('');
    })
    .join('\n');
}

function twipsToCssInches(twips: number | undefined): string | undefined {
  if (!twips || !Number.isFinite(twips) || twips <= 0) {
    return undefined;
  }

  return `${Number((twips / twipsPerInch).toFixed(4))}in`;
}

function pageStyleFromRtfMetrics(metrics: RtfPageMetrics): CSSProperties | undefined {
  const style: CSSProperties = {};
  const width = twipsToCssInches(metrics.widthTwips);
  const height = twipsToCssInches(metrics.heightTwips);
  const marginTop = twipsToCssInches(metrics.marginTopTwips);
  const marginRight = twipsToCssInches(metrics.marginRightTwips);
  const marginBottom = twipsToCssInches(metrics.marginBottomTwips);
  const marginLeft = twipsToCssInches(metrics.marginLeftTwips);

  if (width) {
    style.width = width;
  }

  if (height) {
    style.minHeight = height;
  }

  if (marginTop) {
    style.paddingTop = marginTop;
  }

  if (marginRight) {
    style.paddingRight = marginRight;
  }

  if (marginBottom) {
    style.paddingBottom = marginBottom;
  }

  if (marginLeft) {
    style.paddingLeft = marginLeft;
  }

  if (Object.keys(style).length === 0) {
    return undefined;
  }

  style.boxSizing = 'border-box';
  return style;
}

function parseRtfDocument(rtf: string): ParsedRtfDocument {
  const colors = extractColorTable(rtf);
  const blocks: RtfBlock[] = [];
  const pageMetrics: RtfPageMetrics = {};
  const stateStack: RtfParserState[] = [];
  let state: RtfParserState = {
    skip: false,
    style: {},
    ucSkip: 1,
  };
  let paragraphChildren: RtfInline[] = [];
  let currentTable: RtfTableBlock | null = null;
  let currentRow: RtfTableCell[] | null = null;
  let currentCellChildren: RtfInline[] = [];
  let fieldInstruction = '';
  let pendingFieldHref: string | undefined;
  let unicodeFallbackToSkip = 0;

  function inlineTarget(): RtfInline[] {
    if (currentRow) {
      return currentCellChildren;
    }

    if (currentTable) {
      flushTable();
    }

    return paragraphChildren;
  }

  function appendBreak() {
    if (state.skip) {
      return;
    }

    inlineTarget().push({ type: 'break' });
  }

  function appendText(text: string) {
    if (!text) {
      return;
    }

    if (state.destination === 'fldinst') {
      fieldInstruction += text;
      return;
    }

    if (state.skip) {
      return;
    }

    const target = inlineTarget();
    const href = state.linkHref && isSafeNavigationHref(state.linkHref) ? state.linkHref : undefined;
    const previous = target[target.length - 1];

    if (previous?.type === 'text' && styleKey(previous.style, previous.href) === styleKey(state.style, href)) {
      previous.text += text;
      return;
    }

    target.push({
      type: 'text',
      text,
      style: cloneStyle(state.style),
      href,
    });
  }

  function hasVisibleInlineContent(children: RtfInline[]): boolean {
    return children.some((child) => child.type === 'break' || child.text.trim().length > 0);
  }

  function flushParagraph() {
    if (paragraphChildren.length === 0) {
      return;
    }

    if (hasVisibleInlineContent(paragraphChildren)) {
      blocks.push({ type: 'paragraph', children: paragraphChildren });
    }

    paragraphChildren = [];
  }

  function startTableRow() {
    flushParagraph();

    if (!currentTable) {
      currentTable = { type: 'table', rows: [] };
    }

    currentRow = [];
    currentCellChildren = [];
  }

  function endTableCell() {
    if (!currentRow) {
      return;
    }

    currentRow.push({ children: currentCellChildren });
    currentCellChildren = [];
  }

  function endTableRow() {
    if (!currentTable || !currentRow) {
      return;
    }

    if (currentCellChildren.length > 0 || currentRow.length === 0) {
      endTableCell();
    }

    if (currentRow.length > 0) {
      currentTable.rows.push(currentRow);
    }

    currentRow = null;
    currentCellChildren = [];
  }

  function flushTable() {
    if (!currentTable) {
      return;
    }

    endTableRow();

    if (currentTable.rows.length > 0) {
      blocks.push(currentTable);
    }

    currentTable = null;
    currentRow = null;
    currentCellChildren = [];
  }

  function applyControlWord(word: string, parameter?: number) {
    if (skippedDestinations.has(word)) {
      state.destination = word;
      state.skip = true;
      return;
    }

    if (state.skip) {
      return;
    }

    if (word === 'fldinst') {
      state.destination = 'fldinst';
      fieldInstruction = '';
      return;
    }

    if (word === 'fldrslt') {
      state.destination = 'fldrslt';
      state.linkHref = pendingFieldHref;
      return;
    }

    switch (word) {
      case 'paperw':
        if (parameter && parameter > 0) {
          pageMetrics.widthTwips = parameter;
        }
        break;
      case 'paperh':
        if (parameter && parameter > 0) {
          pageMetrics.heightTwips = parameter;
        }
        break;
      case 'margl':
        if (parameter !== undefined && parameter >= 0) {
          pageMetrics.marginLeftTwips = parameter;
        }
        break;
      case 'margr':
        if (parameter !== undefined && parameter >= 0) {
          pageMetrics.marginRightTwips = parameter;
        }
        break;
      case 'margt':
        if (parameter !== undefined && parameter >= 0) {
          pageMetrics.marginTopTwips = parameter;
        }
        break;
      case 'margb':
        if (parameter !== undefined && parameter >= 0) {
          pageMetrics.marginBottomTwips = parameter;
        }
        break;
      case 'b':
        state.style.bold = parameter !== 0;
        break;
      case 'i':
        state.style.italic = parameter !== 0;
        break;
      case 'ul':
        state.style.underline = parameter !== 0;
        break;
      case 'ulnone':
        state.style.underline = false;
        break;
      case 'strike':
        state.style.strike = parameter !== 0;
        break;
      case 'fs':
        if (parameter && parameter > 0) {
          state.style.fontSize = parameter / 2;
        }
        break;
      case 'cf':
        if (parameter && colors.has(parameter)) {
          state.style.color = colors.get(parameter);
        } else {
          delete state.style.color;
        }
        break;
      case 'cb':
      case 'highlight':
        if (parameter && colors.has(parameter)) {
          state.style.backgroundColor = colors.get(parameter);
        } else {
          delete state.style.backgroundColor;
        }
        break;
      case 'super':
        state.style.verticalAlign = 'super';
        break;
      case 'sub':
        state.style.verticalAlign = 'sub';
        break;
      case 'nosupersub':
        delete state.style.verticalAlign;
        break;
      case 'plain':
        state.style = {};
        break;
      case 'par':
        if (currentRow) {
          appendBreak();
        } else {
          flushParagraph();
        }
        break;
      case 'line':
        appendBreak();
        break;
      case 'tab':
        appendText('\t');
        break;
      case 'bullet':
        appendText('•');
        break;
      case 'emdash':
        appendText('—');
        break;
      case 'endash':
        appendText('–');
        break;
      case 'lquote':
        appendText('‘');
        break;
      case 'rquote':
        appendText('’');
        break;
      case 'ldblquote':
        appendText('“');
        break;
      case 'rdblquote':
        appendText('”');
        break;
      case 'trowd':
        startTableRow();
        break;
      case 'cell':
        endTableCell();
        break;
      case 'row':
        endTableRow();
        break;
      case 'uc':
        if (parameter !== undefined && parameter >= 0) {
          state.ucSkip = parameter;
        }
        break;
      case 'u':
        if (parameter !== undefined) {
          appendText(normalizeUnicodeValue(parameter));
          unicodeFallbackToSkip = state.ucSkip;
        }
        break;
      default:
        break;
    }
  }

  for (let index = 0; index < rtf.length; index += 1) {
    const char = rtf[index];

    if (unicodeFallbackToSkip > 0 && char !== '{' && char !== '}' && char !== '\\') {
      unicodeFallbackToSkip -= 1;
      continue;
    }

    if (char === '{') {
      stateStack.push(cloneState(state));
      continue;
    }

    if (char === '}') {
      if (state.destination === 'fldinst') {
        pendingFieldHref = hrefFromFieldInstruction(fieldInstruction);
      }

      if (state.destination === 'fldrslt') {
        pendingFieldHref = undefined;
      }

      state = stateStack.pop() ?? state;
      continue;
    }

    if (char !== '\\') {
      appendText(char);
      continue;
    }

    const nextChar = rtf[index + 1];

    if (nextChar === undefined) {
      break;
    }

    if (nextChar === '\\' || nextChar === '{' || nextChar === '}') {
      appendText(nextChar);
      index += 1;
      continue;
    }

    if (nextChar === '~') {
      appendText('\u00a0');
      index += 1;
      continue;
    }

    if (nextChar === '-') {
      appendText('\u00ad');
      index += 1;
      continue;
    }

    if (nextChar === '_') {
      appendText('\u2011');
      index += 1;
      continue;
    }

    if (nextChar === '*') {
      index += 1;
      continue;
    }

    if (nextChar === "'") {
      const hex = rtf.slice(index + 2, index + 4);
      const byte = Number.parseInt(hex, 16);

      if (Number.isFinite(byte)) {
        appendText(decodeRtfByte(byte));
        index += 3;
        if (unicodeFallbackToSkip > 0) {
          unicodeFallbackToSkip -= 1;
        }
      }

      continue;
    }

    if (!/[A-Za-z]/.test(nextChar)) {
      index += 1;
      continue;
    }

    let cursor = index + 1;
    let word = '';

    while (cursor < rtf.length && /[A-Za-z]/.test(rtf[cursor])) {
      word += rtf[cursor];
      cursor += 1;
    }

    let parameterText = '';
    let hasParameter = false;

    if (rtf[cursor] === '-') {
      parameterText += '-';
      cursor += 1;
    }

    while (cursor < rtf.length && /\d/.test(rtf[cursor])) {
      hasParameter = true;
      parameterText += rtf[cursor];
      cursor += 1;
    }

    const parameter = hasParameter ? Number(parameterText) : undefined;
    const hadSpaceDelimiter = rtf[cursor] === ' ';

    applyControlWord(word, parameter);
    index = hadSpaceDelimiter ? cursor : cursor - 1;
  }

  flushParagraph();
  flushTable();

  return {
    blocks,
    pageStyle: pageStyleFromRtfMetrics(pageMetrics),
    plainText: plainTextFromBlocks(blocks),
  };
}

function inlineStyleToCss(style: RtfInlineStyle): CSSProperties {
  const textDecoration = [style.underline ? 'underline' : null, style.strike ? 'line-through' : null].filter(Boolean).join(' ');

  return {
    backgroundColor: style.backgroundColor,
    color: style.color,
    fontSize: style.fontSize ? `${style.fontSize}pt` : undefined,
    fontStyle: style.italic ? 'italic' : undefined,
    fontWeight: style.bold ? 700 : undefined,
    textDecoration: textDecoration || undefined,
    verticalAlign: style.verticalAlign,
  };
}

function renderInline(run: RtfInline, searchTerm: string, key: number | string): ReactNode {
  if (run.type === 'break') {
    return <br key={key} />;
  }

  const content = highlightText(run.text, searchTerm);
  const style = inlineStyleToCss(run.style);

  if (run.href && isSafeNavigationHref(run.href)) {
    return (
      <a href={run.href} key={key} rel="noopener noreferrer" style={style} target="_blank">
        {content}
      </a>
    );
  }

  return (
    <span key={key} style={style}>
      {content}
    </span>
  );
}

function renderParagraph(block: RtfParagraphBlock, searchTerm: string, key: number | string): ReactNode {
  return <p key={key}>{block.children.map((child, index) => renderInline(child, searchTerm, index))}</p>;
}

function renderTable(block: RtfTableBlock, searchTerm: string, key: number | string): ReactNode {
  return (
    <table key={key}>
      <tbody>
        {block.rows.map((row, rowIndex) => (
          <tr key={rowIndex}>
            {row.map((cell, cellIndex) => (
              <td key={cellIndex}>{cell.children.map((child, index) => renderInline(child, searchTerm, index))}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function RtfRendererComponent({ file, state, actions }: DocumentRendererProps) {
  const [document, setDocument] = useState<ParsedRtfDocument | null>(null);

  useEffect(() => {
    const decoded = decodeRtfBuffer(file.arrayBuffer);
    const parsedDocument = parseRtfDocument(decoded);

    setDocument(parsedDocument);
    actions.setPageCount(undefined);
    actions.setDocumentInfo({ title: file.fileName });
  }, [actions, file]);

  const matches = useMemo(() => countMatches(document?.plainText ?? '', state.searchTerm), [document, state.searchTerm]);

  useEffect(() => {
    actions.setSearchStats({
      totalMatches: matches,
      currentPageMatches: matches,
      message: state.searchTerm ? `${matches} match${matches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, matches, state.searchTerm]);

  if (!document) {
    return <div className="ldv-renderer-status">Loading RTF...</div>;
  }

  return (
    <article
      className={`ldv-rich-document ldv-rtf-document${document.pageStyle ? ' ldv-paged-rich-document' : ''}`}
      style={{
        ...document.pageStyle,
        fontSize: document.pageStyle ? '16px' : `${16 * state.zoom}px`,
        transform: `rotate(${state.rotation}deg)`,
        transformOrigin: 'center top',
        zoom: document.pageStyle ? state.zoom : undefined,
      }}
    >
      {document.blocks.map((block, index) =>
        block.type === 'table' ? renderTable(block, state.searchTerm, index) : renderParagraph(block, state.searchTerm, index),
      )}
    </article>
  );
}

export const RtfRenderer: DocumentRenderer = {
  id: 'rtf',
  label: 'RTF',
  priority: 31,
  canRender: (file) =>
    file.extension === 'rtf' ||
    file.mimeType === 'application/rtf' ||
    file.mimeType === 'text/rtf' ||
    file.mimeType === 'application/x-rtf',
  Component: RtfRendererComponent,
};
