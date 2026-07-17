import { renderAsync } from 'docx-preview';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { countMatches } from '../utils/highlight';

const blockedTags = new Set(['script', 'iframe', 'object', 'embed', 'link', 'meta']);
const urlAttributes = ['src', 'srcset', 'href', 'poster', 'data'];
const transparentFillValues = new Set(['', 'none', 'transparent', 'inherit']);

function isNetworkReference(value: string): boolean {
  const clean = value.trim();
  return /^https?:\/\//i.test(clean) || /^\/\//.test(clean);
}

function containsNetworkReference(value: string): boolean {
  return /url\(\s*['"]?(?:https?:\/\/|\/\/)/i.test(value) || /@import\s+['"]?(?:https?:\/\/|\/\/)/i.test(value);
}

function normalizeOfficeFontFamily(fontFamily: string): string {
  const clean = fontFamily.trim();

  if (clean.includes(',')) {
    return clean;
  }

  if (/^["']?Arial MT["']?$/i.test(clean)) {
    return '"Arial MT", Arial, Helvetica, sans-serif';
  }

  if (/^["']?Calibri["']?$/i.test(clean)) {
    return 'Calibri, Arial, Helvetica, sans-serif';
  }

  return clean;
}

function normalizeDocxOfficeFontFallbacks(container: HTMLElement): void {
  container.querySelectorAll<HTMLElement>('[style*="font-family"]').forEach((element) => {
    const normalizedFontFamily = normalizeOfficeFontFamily(element.style.fontFamily);

    if (normalizedFontFamily && normalizedFontFamily !== element.style.fontFamily) {
      element.style.fontFamily = normalizedFontFamily;
    }
  });
}

function findInheritedSvgFill(element: Element): string | undefined {
  let current: Element | null = element;

  while (current && current.localName.toLowerCase() !== 'svg') {
    const fill = current.getAttribute('fill')?.trim();

    if (fill && !transparentFillValues.has(fill.toLowerCase())) {
      return fill;
    }

    current = current.parentElement;
  }

  return undefined;
}

function normalizeDocxVmlTextBoxes(container: HTMLElement): void {
  container.querySelectorAll<SVGForeignObjectElement>('svg foreignObject').forEach((foreignObject) => {
    const svg = foreignObject.closest('svg');
    const hasTextContent = Boolean(foreignObject.textContent?.trim());

    if (!svg || !hasTextContent) {
      return;
    }

    const replacement = document.createElement('div');
    const fill = findInheritedSvgFill(foreignObject);

    replacement.className = 'ldv-docx-vml-textbox';
    replacement.setAttribute('style', svg.getAttribute('style') ?? '');

    if (fill) {
      replacement.style.backgroundColor = fill;
    }

    if (svg.style.width) {
      replacement.style.width = svg.style.width;
    }

    if (svg.style.height) {
      replacement.style.height = svg.style.height;
      replacement.style.minHeight = svg.style.height;
    }

    if (Number.parseInt(replacement.style.zIndex, 10) < 0) {
      replacement.style.zIndex = '0';
    }

    replacement.append(...Array.from(foreignObject.childNodes).map((node) => node.cloneNode(true)));
    svg.replaceWith(replacement);
  });
}

function hardenRenderedDocx(container: HTMLElement): string[] {
  const warnings: string[] = [];

  container.querySelectorAll('*').forEach((element) => {
    const tagName = element.tagName.toLowerCase();

    if (blockedTags.has(tagName)) {
      element.remove();
      warnings.push(`Removed unsupported embedded ${tagName} content.`);
      return;
    }

    Array.from(element.attributes).forEach((attribute) => {
      const name = attribute.name.toLowerCase();
      const value = attribute.value;

      if (name.startsWith('on')) {
        element.removeAttribute(attribute.name);
      }

      if (name === 'style' && containsNetworkReference(value)) {
        element.removeAttribute(attribute.name);
        warnings.push('Removed a remote style reference from the DOCX preview.');
      }

      if (urlAttributes.includes(name) && isNetworkReference(value)) {
        element.removeAttribute(attribute.name);
        warnings.push('Removed a remote resource reference from the DOCX preview.');
      }
    });

    if (tagName === 'a') {
      element.setAttribute('rel', 'noopener noreferrer');
      element.setAttribute('target', '_blank');
    }
  });

  return Array.from(new Set(warnings));
}

function pageCountFromContainer(container: HTMLElement): number | undefined {
  const pages = container.querySelectorAll('.docx-wrapper > section.docx, section.docx');
  return pages.length > 0 ? pages.length : undefined;
}

function DocxRendererComponent({ file, state, actions }: DocumentRendererProps) {
  const bodyRef = useRef<HTMLDivElement>(null);
  const styleRef = useRef<HTMLDivElement>(null);
  const [plainText, setPlainText] = useState('');
  const [warnings, setWarnings] = useState<string[]>([]);

  useEffect(() => {
    let cancelled = false;
    const bodyContainer = bodyRef.current;
    const styleContainer = styleRef.current;

    if (!bodyContainer || !styleContainer) {
      return undefined;
    }

    const activeBodyContainer = bodyContainer;
    const activeStyleContainer = styleContainer;

    activeBodyContainer.replaceChildren();
    activeStyleContainer.replaceChildren();
    setPlainText('');
    setWarnings([]);

    async function renderDocx() {
      await renderAsync(file.arrayBuffer.slice(0), activeBodyContainer, activeStyleContainer, {
        breakPages: true,
        className: 'docx',
        experimental: true,
        ignoreFonts: false,
        ignoreHeight: false,
        ignoreLastRenderedPageBreak: false,
        ignoreWidth: false,
        inWrapper: true,
        renderAltChunks: false,
        renderChanges: false,
        renderComments: true,
        renderEndnotes: true,
        renderFooters: true,
        renderFootnotes: true,
        renderHeaders: true,
        trimXmlDeclaration: true,
        useBase64URL: true,
      });

      if (cancelled) {
        return;
      }

      normalizeDocxVmlTextBoxes(activeBodyContainer);
      normalizeDocxOfficeFontFallbacks(activeBodyContainer);

      const nextWarnings = hardenRenderedDocx(activeBodyContainer);
      const nextPageCount = pageCountFromContainer(activeBodyContainer);

      setPlainText(activeBodyContainer.textContent ?? '');
      setWarnings(nextWarnings);
      actions.setPageCount(nextPageCount);
      actions.setDocumentInfo({
        title: file.fileName,
        pageCount: nextPageCount,
        warnings: nextWarnings,
      });
    }

    renderDocx().catch((error: unknown) => {
      if (!cancelled) {
        actions.reportError(error);
      }
    });

    return () => {
      cancelled = true;
      activeBodyContainer.replaceChildren();
      activeStyleContainer.replaceChildren();
    };
  }, [actions, file]);

  const matches = useMemo(() => countMatches(plainText, state.searchTerm), [plainText, state.searchTerm]);

  useEffect(() => {
    actions.setSearchStats({
      totalMatches: matches,
      currentPageMatches: matches,
      message: state.searchTerm ? `${matches} match${matches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, matches, state.searchTerm]);

  useEffect(() => {
    const pages = bodyRef.current?.querySelectorAll<HTMLElement>('.docx-wrapper > section.docx, section.docx');
    const targetPage = pages?.[state.page - 1];

    if (targetPage) {
      targetPage.scrollIntoView({ block: 'start', behavior: 'smooth' });
    }
  }, [state.page]);

  return (
    <div
      className="ldv-docx-preview-shell"
      style={{
        zoom: state.zoom,
        transform: `rotate(${state.rotation}deg)`,
        transformOrigin: 'center top',
      }}
    >
      <div className="ldv-docx-style-container" ref={styleRef} />
      {warnings.length > 0 ? (
        <div className="ldv-warning" role="status">
          Some unsafe embedded DOCX content was removed before preview.
        </div>
      ) : null}
      <div className="ldv-docx-document" ref={bodyRef} />
    </div>
  );
}

export const DocxRenderer: DocumentRenderer = {
  id: 'docx',
  label: 'DOCX',
  priority: 25,
  canRender: (file) =>
    file.extension === 'docx' ||
    file.mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  Component: DocxRendererComponent,
};
