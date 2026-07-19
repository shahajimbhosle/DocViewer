import JSZip from 'jszip';
import { renderAsync } from 'docx-preview';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { RefObject } from 'react';
import type { DocumentRenderer, DocumentRendererProps, FitMode } from '../types';
import { countMatches } from '../utils/highlight';
import { isLocalFragmentHref, isNetworkReference, isSafeNavigationHref } from '../utils/sanitize';

const blockedTags = new Set(['script', 'iframe', 'object', 'embed', 'link', 'meta']);
const urlResourceAttributes = ['src', 'srcset', 'poster', 'data', 'xlink:href'];
const transparentFillValues = new Set(['', 'none', 'transparent', 'inherit']);
const officeSymbolFontPattern = /(?:^|,|\s)(?:symbol|wingdings|wingdings 2|wingdings 3|webdings)(?:,|\s|$)/i;
const officeBulletCharacterMap = new Map([
  ['\uf0b7', '•'],
  ['\uf0a7', '▪'],
  ['\uf0d8', '➢'],
  ['\uf0fc', '✓'],
]);
const emuPerPixel = 9525;
const svgNamespace = 'http://www.w3.org/2000/svg';

interface DocxShapeFallback {
  fill: string;
  heightPx: number;
  kind: 'ellipse' | 'line' | 'rect' | 'roundRect';
  stroke: string;
  strokeWidthPx: number;
  text: string;
  widthPx: number;
}

interface DocxImageFallback {
  alt?: string;
  dataUrl: string;
  heightPx: number;
  kind: 'image';
  widthPx: number;
}

type DocxDrawingFallback = DocxImageFallback | (DocxShapeFallback & { kind: DocxShapeFallback['kind'] });

interface Size {
  width: number;
  height: number;
}

function useElementSize(ref: RefObject<HTMLElement>): Size {
  const [size, setSize] = useState<Size>({ width: 0, height: 0 });

  useEffect(() => {
    const element = ref.current;
    if (!element) {
      return undefined;
    }

    const updateSize = () => {
      setSize({
        width: element.clientWidth,
        height: element.clientHeight,
      });
    };

    updateSize();
    const observer = new ResizeObserver(updateSize);
    observer.observe(element);

    return () => observer.disconnect();
  }, [ref]);

  return size;
}

function measureFirstPage(container: HTMLElement): Size | null {
  const page = container.querySelector<HTMLElement>('.docx-wrapper > section.docx, section.docx');

  if (!page) {
    return null;
  }

  const styles = window.getComputedStyle(page);
  const width = Number.parseFloat(styles.width) || page.offsetWidth;
  const height = Number.parseFloat(styles.height) || page.offsetHeight;

  return width > 0 && height > 0 ? { width, height } : null;
}

function scaleForFitMode(pageSize: Size | null, viewportSize: Size, fitMode: FitMode, zoom: number): number {
  if (!pageSize || fitMode === 'manual') {
    return zoom;
  }

  const availableWidth = Math.max(viewportSize.width - 48, 160);
  const availableHeight = Math.max(viewportSize.height - 48, 160);
  const widthScale = availableWidth / pageSize.width;
  const heightScale = availableHeight / pageSize.height;

  return Math.max(0.1, Math.min(fitMode === 'width' ? widthScale : Math.min(widthScale, heightScale), 6));
}

function containsNetworkReference(value: string): boolean {
  return /url\(\s*['"]?(?:https?:\/\/|\/\/)/i.test(value) || /@import\s+['"]?(?:https?:\/\/|\/\/)/i.test(value);
}

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function childElementsByLocalName(element: Element | undefined, localName: string): Element[] {
  return element ? Array.from(element.children).filter((child) => child.localName === localName) : [];
}

function childElementByLocalName(element: Element | undefined, localName: string): Element | undefined {
  return childElementsByLocalName(element, localName)[0];
}

function elementsByLocalName(root: ParentNode | undefined, localName: string): Element[] {
  if (!root) {
    return [];
  }

  const searchableRoot = root as ParentNode & {
    getElementsByTagNameNS?: (namespace: string, localName: string) => HTMLCollectionOf<Element>;
  };

  if (typeof searchableRoot.getElementsByTagNameNS === 'function') {
    return Array.from(searchableRoot.getElementsByTagNameNS('*', localName));
  }

  return Array.from(root.querySelectorAll('*')).filter((element) => element.localName === localName);
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

function emuToPixels(value: string | null | undefined): number {
  const emu = Number(value ?? 0);
  return Number.isFinite(emu) ? emu / emuPerPixel : 0;
}

function normalizeZipPath(path: string): string {
  const parts: string[] = [];

  path
    .replace(/^\/+/, '')
    .split('/')
    .forEach((part) => {
      if (!part || part === '.') {
        return;
      }

      if (part === '..') {
        parts.pop();
        return;
      }

      parts.push(part);
    });

  return parts.join('/');
}

function resolveZipTarget(sourcePath: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizeZipPath(target);
  }

  const basePath = sourcePath.split('/').slice(0, -1).join('/');
  return normalizeZipPath(`${basePath}/${target}`);
}

function relationshipPathForPart(path: string): string {
  const parts = path.split('/');
  const fileName = parts.pop();

  return `${parts.join('/')}/_rels/${fileName}.rels`;
}

function imageMimeTypeFromPath(path: string): string {
  const extension = path.split('.').pop()?.toLowerCase();

  switch (extension) {
    case 'gif':
      return 'image/gif';
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'svg':
      return 'image/svg+xml';
    case 'webp':
      return 'image/webp';
    case 'tif':
    case 'tiff':
      return 'image/tiff';
    case 'png':
    default:
      return 'image/png';
  }
}

function parseRelationships(relationshipsXml: string | undefined, sourcePath: string): Map<string, { target: string; targetMode?: string }> {
  const relationships = new Map<string, { target: string; targetMode?: string }>();

  if (!relationshipsXml) {
    return relationships;
  }

  elementsByLocalName(parseXml(relationshipsXml), 'Relationship').forEach((relationship) => {
    const id = relationship.getAttribute('Id');
    const target = relationship.getAttribute('Target');

    if (id && target) {
      relationships.set(id, {
        target: resolveZipTarget(sourcePath, target),
        targetMode: relationship.getAttribute('TargetMode') ?? undefined,
      });
    }
  });

  return relationships;
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

function hasOfficeBulletCharacter(value: string): boolean {
  return Array.from(officeBulletCharacterMap.keys()).some((character) => value.includes(character));
}

function normalizeOfficeBulletCharacters(value: string): string {
  let normalized = value;

  officeBulletCharacterMap.forEach((replacement, character) => {
    const code = character.charCodeAt(0).toString(16);
    normalized = normalized
      .split(character)
      .join(replacement)
      .replace(new RegExp(`\\\\0*${code}\\s?`, 'gi'), replacement);
  });

  return normalized;
}

function normalizeDocxNumberingBulletContent(value: string): string {
  const normalized = normalizeOfficeBulletCharacters(value);

  return normalized.replace(/p\.docx-num-[^{]+:before\s*\{[^}]*\}/gi, (rule) => {
    if (!/font-family\s*:\s*(?:"Courier New"|'Courier New'|Courier New)\s*;?/i.test(rule)) {
      return rule;
    }

    return rule.replace(
      /(content\s*:\s*)(["'])o((?:\\9|\\a0|\\00a0|\\0000a0|\s)*)\2/gi,
      (_match, propertyPrefix: string, quote: string, suffix: string) =>
        `${propertyPrefix}${quote}○${suffix}${quote}`,
    );
  });
}

function hasOfficeSymbolFont(fontFamily: string): boolean {
  return officeSymbolFontPattern.test(fontFamily);
}

function normalizeDocxOfficeFontFallbacks(container: HTMLElement): void {
  container.querySelectorAll<HTMLElement>('[style*="font-family"]').forEach((element) => {
    const normalizedFontFamily = normalizeOfficeFontFamily(element.style.fontFamily);

    if (normalizedFontFamily && normalizedFontFamily !== element.style.fontFamily) {
      element.style.fontFamily = normalizedFontFamily;
    }
  });
}

function normalizeDocxBulletGlyphs(bodyContainer: HTMLElement, styleContainer: HTMLElement): void {
  bodyContainer.querySelectorAll<HTMLElement>('style').forEach((styleElement) => {
    styleElement.textContent = normalizeDocxNumberingBulletContent(styleElement.textContent ?? '');
  });
  styleContainer.querySelectorAll<HTMLElement>('style').forEach((styleElement) => {
    styleElement.textContent = normalizeDocxNumberingBulletContent(styleElement.textContent ?? '');
  });

  const walker = document.createTreeWalker(bodyContainer, NodeFilter.SHOW_TEXT);
  let currentNode = walker.nextNode();

  while (currentNode) {
    const textNode = currentNode as Text;
    const currentText = textNode.nodeValue ?? '';

    if (hasOfficeBulletCharacter(currentText)) {
      textNode.nodeValue = normalizeOfficeBulletCharacters(currentText);

      if (textNode.parentElement && hasOfficeSymbolFont(textNode.parentElement.style.fontFamily)) {
        textNode.parentElement.classList.add('ldv-docx-normalized-bullet');
      }
    }

    currentNode = walker.nextNode();
  }
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

    const hasSiblingShapeGeometry = Array.from(
      svg.querySelectorAll('rect, ellipse, line, path, polygon, polyline, circle, image'),
    ).some((element) => !foreignObject.contains(element));

    if (hasSiblingShapeGeometry) {
      svg.classList.add('ldv-docx-vml-graphic');
      foreignObject.classList.add('ldv-docx-vml-textbox-foreign-object');
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

function shapeKindFromPreset(preset: string | null | undefined): DocxShapeFallback['kind'] {
  switch (preset) {
    case 'ellipse':
    case 'oval':
      return 'ellipse';
    case 'line':
    case 'straightConnector1':
      return 'line';
    case 'roundRect':
      return 'roundRect';
    default:
      return 'rect';
  }
}

function colorFromSolidFill(fillElement: Element | undefined, fallback: string): string {
  const srgb = elementsByLocalName(fillElement, 'srgbClr')[0]?.getAttribute('val');

  if (srgb) {
    return `#${srgb}`;
  }

  const scheme = elementsByLocalName(fillElement, 'schemeClr')[0]?.getAttribute('val');

  switch (scheme) {
    case 'accent1':
      return '#4472c4';
    case 'accent2':
      return '#ed7d31';
    case 'accent3':
      return '#a5a5a5';
    case 'accent4':
      return '#ffc000';
    case 'accent5':
      return '#5b9bd5';
    case 'accent6':
      return '#70ad47';
    case 'bg1':
    case 'lt1':
      return '#ffffff';
    case 'dk1':
    case 'tx1':
      return '#000000';
    default:
      return fallback;
  }
}

function textFromShape(shapeElement: Element): string {
  const textBox = elementsByLocalName(shapeElement, 'txbxContent')[0] ?? elementsByLocalName(shapeElement, 'txBody')[0];

  return elementsByLocalName(textBox, 't')
    .map((textNode) => textNode.textContent ?? '')
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function sizeFromDrawing(drawingWrapper: Element, shapeElement: Element | undefined): Size {
  const extent = childElementByLocalName(drawingWrapper, 'extent');
  const spPr = elementsByLocalName(shapeElement, 'spPr')[0];

  return {
    height:
      Math.max(16, emuToPixels(extent?.getAttribute('cy')) || emuToPixels(elementsByLocalName(spPr, 'ext')[0]?.getAttribute('cy')) || 80),
    width:
      Math.max(24, emuToPixels(extent?.getAttribute('cx')) || emuToPixels(elementsByLocalName(spPr, 'ext')[0]?.getAttribute('cx')) || 160),
  };
}

async function parseImageFallbackFromDrawing(
  zip: JSZip,
  drawing: Element,
  relationships: Map<string, { target: string; targetMode?: string }>,
): Promise<DocxImageFallback | undefined> {
  const drawingWrapper = childElementByLocalName(drawing, 'inline') ?? childElementByLocalName(drawing, 'anchor');
  const picture = elementsByLocalName(drawing, 'pic')[0];
  const blip = elementsByLocalName(picture, 'blip')[0];
  const relationshipId = localAttribute(blip, 'embed');

  if (!drawingWrapper || !picture || !relationshipId) {
    return undefined;
  }

  const relationship = relationships.get(relationshipId);

  if (!relationship || relationship.targetMode?.toLowerCase() === 'external' || isNetworkReference(relationship.target)) {
    return undefined;
  }

  const media = zip.file(relationship.target);
  const base64 = await media?.async('base64');

  if (!base64) {
    return undefined;
  }

  const size = sizeFromDrawing(drawingWrapper, picture);
  const metadata = elementsByLocalName(picture, 'cNvPr')[0];

  return {
    alt: metadata?.getAttribute('descr') ?? metadata?.getAttribute('name') ?? undefined,
    dataUrl: `data:${imageMimeTypeFromPath(relationship.target)};base64,${base64}`,
    heightPx: size.height,
    kind: 'image',
    widthPx: size.width,
  };
}

function parseShapeFallbackFromDrawing(drawing: Element): DocxShapeFallback | undefined {
  const drawingWrapper = childElementByLocalName(drawing, 'inline') ?? childElementByLocalName(drawing, 'anchor');
  const graphicData = elementsByLocalName(drawing, 'graphicData')[0];

  if (!drawingWrapper || !graphicData || elementsByLocalName(graphicData, 'pic')[0]) {
    return undefined;
  }

  const shapeElement =
    elementsByLocalName(graphicData, 'wsp')[0] ??
    elementsByLocalName(graphicData, 'sp')[0] ??
    elementsByLocalName(graphicData, 'cxnSp')[0];

  if (!shapeElement) {
    return undefined;
  }

  const size = sizeFromDrawing(drawingWrapper, shapeElement);
  const spPr = elementsByLocalName(shapeElement, 'spPr')[0];
  const line = childElementByLocalName(spPr, 'ln');
  const fillElement = childElementByLocalName(spPr, 'solidFill');
  const noFill = Boolean(childElementByLocalName(spPr, 'noFill'));
  const lineNoFill = Boolean(childElementByLocalName(line, 'noFill'));
  const preset = childElementByLocalName(spPr, 'prstGeom')?.getAttribute('prst');

  return {
    fill: noFill ? 'transparent' : colorFromSolidFill(fillElement, '#e8f3ff'),
    heightPx: size.height,
    kind: shapeKindFromPreset(preset),
    stroke: lineNoFill ? 'transparent' : colorFromSolidFill(childElementByLocalName(line, 'solidFill'), '#2563eb'),
    strokeWidthPx: Math.max(1, Math.min(12, emuToPixels(line?.getAttribute('w')) || 1.5)),
    text: textFromShape(shapeElement),
    widthPx: size.width,
  };
}

async function parseDocxDrawingFallbacks(arrayBuffer: ArrayBuffer): Promise<DocxDrawingFallback[]> {
  try {
    const zip = await JSZip.loadAsync(arrayBuffer.slice(0));
    const documentXml = await zip.file('word/document.xml')?.async('text');

    if (!documentXml) {
      return [];
    }

    const document = parseXml(documentXml);
    const relationships = parseRelationships(await zip.file(relationshipPathForPart('word/document.xml'))?.async('text'), 'word/document.xml');
    const drawings = elementsByLocalName(document, 'drawing');
    const fallbacks = await Promise.all(
      drawings.map(async (drawing) => {
        const imageFallback = await parseImageFallbackFromDrawing(zip, drawing, relationships);
        return imageFallback ?? parseShapeFallbackFromDrawing(drawing);
      }),
    );

    return fallbacks.filter((fallback): fallback is DocxDrawingFallback => Boolean(fallback));
  } catch {
    return [];
  }
}

function appendSvgShape(svg: SVGSVGElement, shape: DocxShapeFallback): void {
  const inset = Math.max(0.5, shape.strokeWidthPx / 2);

  if (shape.kind === 'line') {
    const line = document.createElementNS(svgNamespace, 'line');
    line.setAttribute('x1', String(inset));
    line.setAttribute('y1', String(shape.heightPx - inset));
    line.setAttribute('x2', String(shape.widthPx - inset));
    line.setAttribute('y2', String(inset));
    line.setAttribute('stroke', shape.stroke);
    line.setAttribute('stroke-linecap', 'round');
    line.setAttribute('stroke-width', String(shape.strokeWidthPx));
    svg.append(line);
    return;
  }

  const element =
    shape.kind === 'ellipse' ? document.createElementNS(svgNamespace, 'ellipse') : document.createElementNS(svgNamespace, 'rect');

  if (shape.kind === 'ellipse') {
    element.setAttribute('cx', String(shape.widthPx / 2));
    element.setAttribute('cy', String(shape.heightPx / 2));
    element.setAttribute('rx', String(Math.max(1, shape.widthPx / 2 - inset)));
    element.setAttribute('ry', String(Math.max(1, shape.heightPx / 2 - inset)));
  } else {
    element.setAttribute('x', String(inset));
    element.setAttribute('y', String(inset));
    element.setAttribute('width', String(Math.max(1, shape.widthPx - inset * 2)));
    element.setAttribute('height', String(Math.max(1, shape.heightPx - inset * 2)));
    if (shape.kind === 'roundRect') {
      element.setAttribute('rx', String(Math.min(shape.widthPx, shape.heightPx) * 0.12));
    }
  }

  element.setAttribute('fill', shape.fill);
  element.setAttribute('stroke', shape.stroke);
  element.setAttribute('stroke-width', String(shape.strokeWidthPx));
  svg.append(element);
}

function createDocxShapeFallbackElement(shape: DocxShapeFallback): HTMLElement {
  const wrapper = document.createElement('span');
  const svg = document.createElementNS(svgNamespace, 'svg');

  wrapper.className = 'ldv-docx-shape-fallback';
  wrapper.style.width = `${shape.widthPx}px`;
  wrapper.style.height = `${shape.heightPx}px`;
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('viewBox', `0 0 ${shape.widthPx} ${shape.heightPx}`);
  appendSvgShape(svg, shape);
  wrapper.append(svg);

  if (shape.text) {
    const label = document.createElement('span');
    label.className = 'ldv-docx-shape-fallback-text';
    label.textContent = shape.text;
    wrapper.append(label);
  }

  return wrapper;
}

function createDocxImageFallbackElement(image: DocxImageFallback): HTMLElement {
  const wrapper = document.createElement('span');
  const element = document.createElement('img');

  wrapper.className = 'ldv-docx-image-fallback';
  wrapper.style.width = `${image.widthPx}px`;
  wrapper.style.height = `${image.heightPx}px`;
  element.alt = image.alt ?? 'DOCX image';
  element.src = image.dataUrl;
  wrapper.append(element);

  return wrapper;
}

function createDocxDrawingFallbackElement(fallback: DocxDrawingFallback): HTMLElement {
  return fallback.kind === 'image' ? createDocxImageFallbackElement(fallback) : createDocxShapeFallbackElement(fallback);
}

function hasRenderedDrawingContent(element: HTMLElement): boolean {
  const hasResolvedImage = Array.from(element.querySelectorAll<HTMLImageElement>('img')).some((image) =>
    Boolean(image.getAttribute('src')?.trim()),
  );
  const hasSvgContent = Array.from(element.querySelectorAll<SVGSVGElement>('svg')).some((svg) => svg.children.length > 0);

  return hasResolvedImage || hasSvgContent || Boolean(element.querySelector('canvas')) || Boolean(element.textContent?.trim());
}

function isPotentialDrawingWrapper(element: HTMLElement): boolean {
  const style = element.getAttribute('style') ?? '';
  return style.includes('inline-block') && style.includes('position: relative');
}

function applyDocxDrawingFallbacks(container: HTMLElement, fallbacks: DocxDrawingFallback[]): void {
  if (fallbacks.length === 0) {
    return;
  }

  const wrappers = Array.from(container.querySelectorAll<HTMLElement>('div')).filter(isPotentialDrawingWrapper);
  const appendTarget = container.querySelector<HTMLElement>('.docx-wrapper > section.docx, section.docx') ?? container;

  fallbacks.forEach((fallback, index) => {
    const target = wrappers[index];

    if (target && !hasRenderedDrawingContent(target)) {
      target.style.width ||= `${fallback.widthPx}px`;
      target.style.height ||= `${fallback.heightPx}px`;
      target.replaceChildren(createDocxDrawingFallbackElement(fallback));
      return;
    }

    if (!target) {
      appendTarget.append(createDocxDrawingFallbackElement(fallback));
    }
  });
}

function decodeFragmentId(fragment: string): string {
  try {
    return decodeURIComponent(fragment);
  } catch {
    return fragment;
  }
}

function findLocalFragmentTarget(container: HTMLElement, href: string): HTMLElement | undefined {
  const fragment = href.trim().slice(1);
  const decodedFragment = decodeFragmentId(fragment);

  if (!fragment) {
    return undefined;
  }

  return Array.from(container.querySelectorAll<HTMLElement>('[id], a[name]')).find((element) => {
    const name = element.getAttribute('name');
    return element.id === fragment || element.id === decodedFragment || name === fragment || name === decodedFragment;
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

      if (urlResourceAttributes.includes(name) && isNetworkReference(value)) {
        element.removeAttribute(attribute.name);
        warnings.push('Removed a remote resource reference from the DOCX preview.');
      }

      if (name === 'href' && tagName === 'a' && !isSafeNavigationHref(value)) {
        element.removeAttribute(attribute.name);
        warnings.push('Removed an unsafe link from the DOCX preview.');
      }

      if (name === 'href' && tagName !== 'a' && isNetworkReference(value)) {
        element.removeAttribute(attribute.name);
        warnings.push('Removed a remote resource reference from the DOCX preview.');
      }
    });

    if (tagName === 'a') {
      const href = element.getAttribute('href');

      if (href && !isLocalFragmentHref(href)) {
        element.setAttribute('rel', 'noopener noreferrer');
        element.setAttribute('target', '_blank');
      } else {
        element.removeAttribute('rel');
        element.removeAttribute('target');
      }
    }
  });

  return Array.from(new Set(warnings));
}

function pageCountFromContainer(container: HTMLElement): number | undefined {
  const pages = container.querySelectorAll('.docx-wrapper > section.docx, section.docx');
  return pages.length > 0 ? pages.length : undefined;
}

function DocxRendererComponent({ file, state, actions, viewportRef }: DocumentRendererProps) {
  const bodyRef = useRef<HTMLDivElement>(null);
  const styleRef = useRef<HTMLDivElement>(null);
  const viewportSize = useElementSize(viewportRef);
  const [plainText, setPlainText] = useState('');
  const [pageSize, setPageSize] = useState<Size | null>(null);
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
    setPageSize(null);
    setWarnings([]);
    actions.setLoading(true);

    async function renderDocx() {
      const [drawingFallbacks] = await Promise.all([
        parseDocxDrawingFallbacks(file.arrayBuffer),
        renderAsync(file.arrayBuffer.slice(0), activeBodyContainer, activeStyleContainer, {
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
        }),
      ]);

      if (cancelled) {
        return;
      }

      normalizeDocxVmlTextBoxes(activeBodyContainer);
      normalizeDocxOfficeFontFallbacks(activeBodyContainer);
      normalizeDocxBulletGlyphs(activeBodyContainer, activeStyleContainer);
      applyDocxDrawingFallbacks(activeBodyContainer, drawingFallbacks);

      const nextWarnings = hardenRenderedDocx(activeBodyContainer);
      const nextPageCount = pageCountFromContainer(activeBodyContainer);

      setPageSize(measureFirstPage(activeBodyContainer));
      setPlainText(activeBodyContainer.textContent ?? '');
      setWarnings(nextWarnings);
      actions.setPageCount(nextPageCount);
      actions.setDocumentInfo({
        title: file.fileName,
        pageCount: nextPageCount,
        warnings: nextWarnings,
      });
      actions.setLoading(false);
    }

    renderDocx().catch((error: unknown) => {
      if (!cancelled) {
        actions.setLoading(false);
        actions.reportError(error);
      }
    });

    return () => {
      cancelled = true;
      actions.setLoading(false);
      activeBodyContainer.replaceChildren();
      activeStyleContainer.replaceChildren();
    };
  }, [actions, file]);

  const matches = useMemo(() => countMatches(plainText, state.searchTerm), [plainText, state.searchTerm]);
  const displayScale = useMemo(
    () => scaleForFitMode(pageSize, viewportSize, state.fitMode, state.zoom),
    [pageSize, state.fitMode, state.zoom, viewportSize.height, viewportSize.width],
  );

  useEffect(() => {
    actions.setSearchStats({
      totalMatches: matches,
      currentPageMatches: matches,
      message: state.searchTerm ? `${matches} match${matches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, matches, state.searchTerm]);

  useEffect(() => {
    const container = bodyRef.current;

    if (!container) {
      return undefined;
    }

    const activeContainer = container;

    function handleClick(event: MouseEvent) {
      const target = event.target instanceof Element ? event.target.closest<HTMLAnchorElement>('a[href]') : null;
      const href = target?.getAttribute('href') ?? '';

      if (!target || !isLocalFragmentHref(href)) {
        return;
      }

      event.preventDefault();
      findLocalFragmentTarget(activeContainer, href)?.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }

    activeContainer.addEventListener('click', handleClick);

    return () => {
      activeContainer.removeEventListener('click', handleClick);
    };
  }, []);

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
        zoom: displayScale,
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
