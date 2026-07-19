import JSZip from 'jszip';
import { useEffect, useMemo, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { countMatches, highlightText } from '../utils/highlight';
import { inferMimeType } from '../utils/mime';
import { isNetworkReference, isSafeNavigationHref } from '../utils/sanitize';

interface OdtStyleRule {
  family?: string;
  parentName?: string;
  paragraphStyle?: CSSProperties;
  textStyle?: CSSProperties;
  tableCellStyle?: CSSProperties;
  tableStyle?: CSSProperties;
}

interface ParsedOdtDocument {
  imageSources: Map<string, string>;
  pageStyle?: CSSProperties;
  plainText: string;
  root: Element;
  styles: Map<string, OdtStyleRule>;
}

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function localAttribute(element: Element | undefined, localName: string): string | null {
  if (!element) {
    return null;
  }

  const directValue = element.getAttribute(localName);
  if (directValue !== null) {
    return directValue;
  }

  return Array.from(element.attributes).find((attribute) => attribute.localName === localName)?.value ?? null;
}

function childElementsByLocalName(element: Element, localName: string): Element[] {
  return Array.from(element.children).filter((child) => child.localName === localName);
}

function childElementByLocalName(element: Element, localName: string): Element | undefined {
  return childElementsByLocalName(element, localName)[0];
}

function elementsByLocalName(root: ParentNode, localName: string): Element[] {
  const searchableRoot = root as ParentNode & {
    getElementsByTagNameNS?: (namespace: string, localName: string) => HTMLCollectionOf<Element>;
  };

  if (typeof searchableRoot.getElementsByTagNameNS === 'function') {
    return Array.from(searchableRoot.getElementsByTagNameNS('*', localName));
  }

  return Array.from(root.querySelectorAll('*')).filter((element) => element.localName === localName);
}

function numberAttribute(element: Element, localName: string, fallback = 1): number {
  const value = Number(localAttribute(element, localName) ?? fallback);
  return Number.isFinite(value) && value > 0 ? Math.floor(value) : fallback;
}

function isVerticalWritingMode(value: string | null): boolean {
  const writingMode = value?.trim().toLowerCase();
  return writingMode === 'tb-rl' || writingMode === 'tb-lr' || writingMode === 'vertical-rl' || writingMode === 'vertical-lr';
}

function isRightToLeftWritingMode(value: string | null): boolean {
  const writingMode = value?.trim().toLowerCase();
  return writingMode === 'rl-tb' || writingMode === 'tb-rl' || writingMode === 'vertical-rl';
}

function styleFromParagraphProperties(element: Element | undefined): CSSProperties {
  const style: CSSProperties = {};
  const textAlign = localAttribute(element, 'text-align');
  const lineHeight = localAttribute(element, 'line-height');
  const marginTop = localAttribute(element, 'margin-top');
  const marginRight = localAttribute(element, 'margin-right');
  const marginBottom = localAttribute(element, 'margin-bottom');
  const marginLeft = localAttribute(element, 'margin-left');
  const backgroundColor = localAttribute(element, 'background-color');
  const writingMode = localAttribute(element, 'writing-mode');

  if (textAlign) {
    style.textAlign = textAlign as CSSProperties['textAlign'];
  }

  if (lineHeight) {
    style.lineHeight = lineHeight;
  }

  if (marginTop) {
    style.marginTop = marginTop;
  }

  if (marginRight) {
    style.marginRight = marginRight;
  }

  if (marginBottom) {
    style.marginBottom = marginBottom;
  }

  if (marginLeft) {
    style.marginLeft = marginLeft;
  }

  if (backgroundColor && backgroundColor !== 'transparent') {
    style.backgroundColor = backgroundColor;
  }

  if (isVerticalWritingMode(writingMode)) {
    style.writingMode = 'vertical-rl';
  }

  if (isRightToLeftWritingMode(writingMode)) {
    style.direction = 'rtl';
  }

  return style;
}

function styleFromTextProperties(element: Element | undefined): CSSProperties {
  const style: CSSProperties = {};
  const color = localAttribute(element, 'color');
  const backgroundColor = localAttribute(element, 'background-color');
  const fontSize = localAttribute(element, 'font-size');
  const fontFamily = localAttribute(element, 'font-family') ?? localAttribute(element, 'font-name');
  const fontWeight = localAttribute(element, 'font-weight');
  const fontStyle = localAttribute(element, 'font-style');
  const underlineStyle = localAttribute(element, 'text-underline-style');
  const textLineThroughStyle = localAttribute(element, 'text-line-through-style');

  if (color) {
    style.color = color;
  }

  if (backgroundColor && backgroundColor !== 'transparent') {
    style.backgroundColor = backgroundColor;
  }

  if (fontSize) {
    style.fontSize = fontSize;
  }

  if (fontFamily) {
    style.fontFamily = fontFamily;
  }

  if (fontWeight) {
    style.fontWeight = fontWeight === 'bold' ? 700 : fontWeight;
  }

  if (fontStyle) {
    style.fontStyle = fontStyle as CSSProperties['fontStyle'];
  }

  if (underlineStyle && underlineStyle !== 'none') {
    style.textDecorationLine = 'underline';
  }

  if (textLineThroughStyle && textLineThroughStyle !== 'none') {
    style.textDecorationLine = style.textDecorationLine ? `${style.textDecorationLine} line-through` : 'line-through';
  }

  return style;
}

function styleFromTableCellProperties(element: Element | undefined): CSSProperties {
  const style: CSSProperties = {};
  const backgroundColor = localAttribute(element, 'background-color');
  const border = localAttribute(element, 'border');
  const borderTop = localAttribute(element, 'border-top');
  const borderRight = localAttribute(element, 'border-right');
  const borderBottom = localAttribute(element, 'border-bottom');
  const borderLeft = localAttribute(element, 'border-left');
  const verticalAlign = localAttribute(element, 'vertical-align');

  if (backgroundColor && backgroundColor !== 'transparent') {
    style.backgroundColor = backgroundColor;
  }

  if (border && border !== 'none') {
    style.border = border;
  }

  if (borderTop && borderTop !== 'none') {
    style.borderTop = borderTop;
  }

  if (borderRight && borderRight !== 'none') {
    style.borderRight = borderRight;
  }

  if (borderBottom && borderBottom !== 'none') {
    style.borderBottom = borderBottom;
  }

  if (borderLeft && borderLeft !== 'none') {
    style.borderLeft = borderLeft;
  }

  if (verticalAlign) {
    style.verticalAlign = verticalAlign === 'middle' ? 'middle' : (verticalAlign as CSSProperties['verticalAlign']);
  }

  return style;
}

function styleFromTableProperties(element: Element | undefined): CSSProperties {
  const style: CSSProperties = {};
  const width = localAttribute(element, 'width');
  const align = localAttribute(element, 'align');

  if (width) {
    style.width = width;
  }

  if (align === 'center') {
    style.marginLeft = 'auto';
    style.marginRight = 'auto';
  }

  if (align === 'right' || align === 'end') {
    style.marginLeft = 'auto';
  }

  return style;
}

function styleFromPageLayoutProperties(element: Element | undefined): CSSProperties | undefined {
  const style: CSSProperties = {};
  const width = localAttribute(element, 'page-width');
  const height = localAttribute(element, 'page-height');
  const marginTop = localAttribute(element, 'margin-top') ?? localAttribute(element, 'margin');
  const marginRight = localAttribute(element, 'margin-right') ?? localAttribute(element, 'margin');
  const marginBottom = localAttribute(element, 'margin-bottom') ?? localAttribute(element, 'margin');
  const marginLeft = localAttribute(element, 'margin-left') ?? localAttribute(element, 'margin');

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

function parseStyleRule(styleElement: Element): OdtStyleRule {
  const paragraphProperties = childElementByLocalName(styleElement, 'paragraph-properties');
  const textProperties = childElementByLocalName(styleElement, 'text-properties');
  const tableCellProperties = childElementByLocalName(styleElement, 'table-cell-properties');
  const tableProperties = childElementByLocalName(styleElement, 'table-properties');

  return {
    family: localAttribute(styleElement, 'family') ?? undefined,
    parentName: localAttribute(styleElement, 'parent-style-name') ?? undefined,
    paragraphStyle: styleFromParagraphProperties(paragraphProperties),
    textStyle: styleFromTextProperties(textProperties),
    tableCellStyle: styleFromTableCellProperties(tableCellProperties),
    tableStyle: styleFromTableProperties(tableProperties),
  };
}

function mergeCssProperties(...styles: Array<CSSProperties | undefined>): CSSProperties | undefined {
  const merged = Object.assign({}, ...styles.filter(Boolean));
  return Object.keys(merged).length > 0 ? merged : undefined;
}

function mergeStyleRules(parent: OdtStyleRule | undefined, child: OdtStyleRule): OdtStyleRule {
  return {
    family: child.family ?? parent?.family,
    parentName: child.parentName,
    paragraphStyle: mergeCssProperties(parent?.paragraphStyle, child.paragraphStyle),
    textStyle: mergeCssProperties(parent?.textStyle, child.textStyle),
    tableCellStyle: mergeCssProperties(parent?.tableCellStyle, child.tableCellStyle),
    tableStyle: mergeCssProperties(parent?.tableStyle, child.tableStyle),
  };
}

function parseStyleMap(...xmlDocuments: Array<Document | undefined>): Map<string, OdtStyleRule> {
  const rawStyles = new Map<string, OdtStyleRule>();
  const resolvedStyles = new Map<string, OdtStyleRule>();

  xmlDocuments.forEach((xmlDocument) => {
    if (!xmlDocument) {
      return;
    }

    elementsByLocalName(xmlDocument, 'style').forEach((styleElement) => {
      const name = localAttribute(styleElement, 'name');

      if (name) {
        rawStyles.set(name, parseStyleRule(styleElement));
      }
    });
  });

  function resolve(name: string, stack = new Set<string>()): OdtStyleRule | undefined {
    const cached = resolvedStyles.get(name);

    if (cached) {
      return cached;
    }

    const raw = rawStyles.get(name);

    if (!raw || stack.has(name)) {
      return raw;
    }

    stack.add(name);
    const parent = raw.parentName ? resolve(raw.parentName, stack) : undefined;
    const resolved = mergeStyleRules(parent, raw);
    resolvedStyles.set(name, resolved);
    stack.delete(name);
    return resolved;
  }

  rawStyles.forEach((_, name) => resolve(name));
  return resolvedStyles;
}

function parsePageStyle(...xmlDocuments: Array<Document | undefined>): CSSProperties | undefined {
  const pageLayouts = new Map<string, CSSProperties>();

  xmlDocuments.forEach((xmlDocument) => {
    if (!xmlDocument) {
      return;
    }

    elementsByLocalName(xmlDocument, 'page-layout').forEach((pageLayoutElement) => {
      const name = localAttribute(pageLayoutElement, 'name');
      const pageLayoutStyle = styleFromPageLayoutProperties(childElementByLocalName(pageLayoutElement, 'page-layout-properties'));

      if (name && pageLayoutStyle) {
        pageLayouts.set(name, pageLayoutStyle);
      }
    });
  });

  for (const xmlDocument of xmlDocuments) {
    if (!xmlDocument) {
      continue;
    }

    for (const masterPage of elementsByLocalName(xmlDocument, 'master-page')) {
      const pageLayoutName = localAttribute(masterPage, 'page-layout-name');
      const pageStyle = pageLayoutName ? pageLayouts.get(pageLayoutName) : undefined;

      if (pageStyle) {
        return pageStyle;
      }
    }
  }

  return pageLayouts.values().next().value;
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\.?\//, '');
}

async function imageSourcesFromZip(zip: JSZip, contentDocument: Document): Promise<Map<string, string>> {
  const imageSources = new Map<string, string>();
  const imageHrefs = Array.from(
    new Set(
      elementsByLocalName(contentDocument, 'image')
        .map((image) => localAttribute(image, 'href'))
        .filter((href): href is string => Boolean(href && !isNetworkReference(href))),
    ),
  );

  await Promise.all(
    imageHrefs.map(async (href) => {
      const zipPath = normalizeZipPath(href);
      const file = zip.file(zipPath);

      if (!file) {
        return;
      }

      const base64 = await file.async('base64');
      imageSources.set(href, `data:${inferMimeType(zipPath)};base64,${base64}`);
    }),
  );

  return imageSources;
}

function textFromNode(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent ?? '';
  }

  if (!(node instanceof Element)) {
    return node.textContent ?? '';
  }

  if (node.localName === 'line-break') {
    return '\n';
  }

  if (node.localName === 'tab') {
    return '\t';
  }

  if (node.localName === 's') {
    return ' '.repeat(numberAttribute(node, 'c', 1));
  }

  const childText = Array.from(node.childNodes).map(textFromNode).join('');

  if (['h', 'p', 'list-item', 'table-row'].includes(node.localName)) {
    return `${childText}\n`;
  }

  return childText;
}

function styleForName(styles: Map<string, OdtStyleRule>, styleName: string | null): OdtStyleRule | undefined {
  return styleName ? styles.get(styleName) : undefined;
}

function renderInlineChildren(element: Element, document: ParsedOdtDocument, searchTerm: string): ReactNode[] {
  return Array.from(element.childNodes).map((node, index) => renderInlineNode(node, document, searchTerm, index));
}

function renderInlineNode(node: Node, document: ParsedOdtDocument, searchTerm: string, key: number | string): ReactNode {
  if (node.nodeType === Node.TEXT_NODE) {
    return <span key={key}>{highlightText(node.textContent ?? '', searchTerm)}</span>;
  }

  if (!(node instanceof Element)) {
    return <span key={key}>{highlightText(node.textContent ?? '', searchTerm)}</span>;
  }

  if (node.localName === 'line-break') {
    return <br key={key} />;
  }

  if (node.localName === 'tab') {
    return <span key={key}>{'\t'}</span>;
  }

  if (node.localName === 's') {
    return <span key={key}>{' '.repeat(numberAttribute(node, 'c', 1))}</span>;
  }

  if (node.localName === 'span') {
    const style = styleForName(document.styles, localAttribute(node, 'style-name'));
    return (
      <span key={key} style={style?.textStyle}>
        {renderInlineChildren(node, document, searchTerm)}
      </span>
    );
  }

  if (node.localName === 'a') {
    const href = localAttribute(node, 'href');
    const safeHref = href && isSafeNavigationHref(href) ? href : undefined;

    return (
      <a href={safeHref} key={key} rel="noopener noreferrer" target="_blank">
        {renderInlineChildren(node, document, searchTerm)}
      </a>
    );
  }

  if (node.localName === 'note') {
    const citation = childElementByLocalName(node, 'note-citation')?.textContent?.trim();
    const body = childElementByLocalName(node, 'note-body');

    return (
      <span className="ldv-odt-note" key={key}>
        {citation ? <sup>{citation}</sup> : null}
        {body ? <span className="ldv-odt-note-body">{textFromNode(body).trim()}</span> : null}
      </span>
    );
  }

  if (node.localName === 'frame' || node.localName === 'image') {
    return renderImageNode(node, document, key);
  }

  return <span key={key}>{renderInlineChildren(node, document, searchTerm)}</span>;
}

function renderImageNode(element: Element, document: ParsedOdtDocument, key: number | string): ReactNode {
  const imageElement = element.localName === 'image' ? element : elementsByLocalName(element, 'image')[0];
  const href = imageElement ? localAttribute(imageElement, 'href') : undefined;
  const source = href ? document.imageSources.get(href) : undefined;
  const title = childElementByLocalName(element, 'title')?.textContent?.trim();
  const description = childElementByLocalName(element, 'desc')?.textContent?.trim();
  const width = localAttribute(element, 'width');
  const height = localAttribute(element, 'height');

  if (!source) {
    return null;
  }

  return (
    <img
      alt={description || title || ''}
      className="ldv-odt-image"
      key={key}
      src={source}
      style={{
        height: height || undefined,
        width: width || undefined,
      }}
    />
  );
}

function renderParagraph(element: Element, document: ParsedOdtDocument, searchTerm: string, key: number | string): ReactNode {
  const style = styleForName(document.styles, localAttribute(element, 'style-name'));
  return (
    <p key={key} style={mergeCssProperties(style?.paragraphStyle, style?.textStyle)}>
      {renderInlineChildren(element, document, searchTerm)}
    </p>
  );
}

function renderHeading(element: Element, document: ParsedOdtDocument, searchTerm: string, key: number | string): ReactNode {
  const level = Math.max(1, Math.min(numberAttribute(element, 'outline-level', 1), 6));
  const Tag = `h${level}` as keyof JSX.IntrinsicElements;
  const style = styleForName(document.styles, localAttribute(element, 'style-name'));

  return (
    <Tag key={key} style={mergeCssProperties(style?.paragraphStyle, style?.textStyle)}>
      {renderInlineChildren(element, document, searchTerm)}
    </Tag>
  );
}

function renderList(element: Element, document: ParsedOdtDocument, searchTerm: string, key: number | string): ReactNode {
  return (
    <ul className="ldv-odt-list" key={key}>
      {childElementsByLocalName(element, 'list-item').map((item, itemIndex) => (
        <li key={itemIndex}>{renderBlockChildren(item, document, searchTerm)}</li>
      ))}
    </ul>
  );
}

function renderTable(element: Element, document: ParsedOdtDocument, searchTerm: string, key: number | string): ReactNode {
  const style = styleForName(document.styles, localAttribute(element, 'style-name'));

  return (
    <table className="ldv-odt-table" key={key} style={style?.tableStyle}>
      <tbody>
        {childElementsByLocalName(element, 'table-row').map((row, rowIndex) => (
          <tr key={rowIndex}>
            {Array.from(row.children)
              .filter((child) => child.localName === 'table-cell' || child.localName === 'covered-table-cell')
              .map((cell, cellIndex) => {
                if (cell.localName === 'covered-table-cell') {
                  return null;
                }

                const cellStyle = styleForName(document.styles, localAttribute(cell, 'style-name'));
                const colSpan = numberAttribute(cell, 'number-columns-spanned', 1);
                const rowSpan = numberAttribute(cell, 'number-rows-spanned', 1);

                return (
                  <td
                    colSpan={colSpan > 1 ? colSpan : undefined}
                    key={cellIndex}
                    rowSpan={rowSpan > 1 ? rowSpan : undefined}
                    style={cellStyle?.tableCellStyle}
                  >
                    {renderBlockChildren(cell, document, searchTerm)}
                  </td>
                );
              })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function renderBlockNode(node: Node, document: ParsedOdtDocument, searchTerm: string, key: number | string): ReactNode {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent ?? '';
    return text.trim() ? <p key={key}>{highlightText(text, searchTerm)}</p> : null;
  }

  if (!(node instanceof Element)) {
    return null;
  }

  if (node.localName === 'h') {
    return renderHeading(node, document, searchTerm, key);
  }

  if (node.localName === 'p') {
    return renderParagraph(node, document, searchTerm, key);
  }

  if (node.localName === 'list') {
    return renderList(node, document, searchTerm, key);
  }

  if (node.localName === 'table') {
    return renderTable(node, document, searchTerm, key);
  }

  if (node.localName === 'section' || node.localName === 'text') {
    return <section key={key}>{renderBlockChildren(node, document, searchTerm)}</section>;
  }

  if (node.localName === 'soft-page-break') {
    return <hr className="ldv-odt-page-break" key={key} />;
  }

  if (node.localName === 'frame' || node.localName === 'image') {
    return renderImageNode(node, document, key);
  }

  return <div key={key}>{renderBlockChildren(node, document, searchTerm)}</div>;
}

function renderBlockChildren(element: Element, document: ParsedOdtDocument, searchTerm: string): ReactNode[] {
  return Array.from(element.childNodes).map((node, index) => renderBlockNode(node, document, searchTerm, index));
}

async function parseOdtDocument(arrayBuffer: ArrayBuffer): Promise<ParsedOdtDocument> {
  const zip = await JSZip.loadAsync(arrayBuffer.slice(0));
  const contentXml = await zip.file('content.xml')?.async('text');

  if (!contentXml) {
    throw new Error('Unable to find content.xml in this ODT file.');
  }

  const stylesXml = await zip.file('styles.xml')?.async('text');
  const contentDocument = parseXml(contentXml);
  const stylesDocument = stylesXml ? parseXml(stylesXml) : undefined;
  const textRoot = elementsByLocalName(contentDocument, 'text')[0];

  if (!textRoot) {
    throw new Error('Unable to find text content in this ODT file.');
  }

  return {
    imageSources: await imageSourcesFromZip(zip, contentDocument),
    pageStyle: parsePageStyle(stylesDocument, contentDocument),
    plainText: textFromNode(textRoot),
    root: textRoot,
    styles: parseStyleMap(stylesDocument, contentDocument),
  };
}

function OdtRendererComponent({ file, state, actions }: DocumentRendererProps) {
  const [document, setDocument] = useState<ParsedOdtDocument | null>(null);

  useEffect(() => {
    let cancelled = false;
    setDocument(null);
    actions.setLoading(true);

    parseOdtDocument(file.arrayBuffer).then(
      (parsedDocument) => {
        if (cancelled) {
          return;
        }

        setDocument(parsedDocument);
        actions.setPageCount(undefined);
        actions.setDocumentInfo({ title: file.fileName });
        actions.setLoading(false);
      },
      (error: unknown) => {
        if (!cancelled) {
          actions.setLoading(false);
          actions.reportError(error);
        }
      },
    );

    return () => {
      cancelled = true;
      actions.setLoading(false);
    };
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
    return <div className="ldv-renderer-status">Loading ODT...</div>;
  }

  return (
    <article
      className={`ldv-rich-document ldv-odt-document${document.pageStyle ? ' ldv-paged-rich-document' : ''}`}
      style={{
        ...document.pageStyle,
        fontSize: document.pageStyle ? '16px' : `${16 * state.zoom}px`,
        transform: `rotate(${state.rotation}deg)`,
        transformOrigin: 'center top',
        zoom: document.pageStyle ? state.zoom : undefined,
      }}
    >
      {renderBlockChildren(document.root, document, state.searchTerm)}
    </article>
  );
}

export const OdtRenderer: DocumentRenderer = {
  id: 'odt',
  label: 'ODT',
  priority: 32,
  canRender: (file) => file.extension === 'odt' || file.mimeType === 'application/vnd.oasis.opendocument.text',
  Component: OdtRendererComponent,
};
