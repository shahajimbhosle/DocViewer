import DOMPurify from 'dompurify';

const networkResourceAttributes = ['src', 'srcset', 'poster', 'data', 'xlink:href'];

export function isNetworkReference(value: string): boolean {
  const clean = value.trim();
  return /^https?:\/\//i.test(clean) || /^\/\//.test(clean);
}

export function isLocalFragmentHref(value: string): boolean {
  return value.trim().startsWith('#');
}

export function isSafeNavigationHref(value: string): boolean {
  const clean = value.trim();

  if (!clean) {
    return false;
  }

  const compact = clean.replace(/[\u0000-\u001F\u007F\s]+/g, '').toLowerCase();

  if (compact.startsWith('javascript:') || compact.startsWith('vbscript:') || compact.startsWith('data:')) {
    return false;
  }

  return true;
}

export function sanitizeHtml(html: string): string {
  const sanitized = DOMPurify.sanitize(html, {
    FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'link', 'meta'],
    FORBID_ATTR: ['style'],
  });

  const template = document.createElement('template');
  template.innerHTML = sanitized;

  template.content.querySelectorAll('*').forEach((element) => {
    const tagName = element.tagName.toLowerCase();

    Array.from(element.attributes).forEach((attribute) => {
      if (attribute.name.toLowerCase().startsWith('on')) {
        element.removeAttribute(attribute.name);
      }
    });

    networkResourceAttributes.forEach((attribute) => {
      const value = element.getAttribute(attribute);
      if (value && isNetworkReference(value)) {
        element.removeAttribute(attribute);
      }
    });

    if (tagName !== 'a') {
      const href = element.getAttribute('href');
      if (href && isNetworkReference(href)) {
        element.removeAttribute('href');
      }
    }

    if (tagName === 'a') {
      const href = element.getAttribute('href');

      if (href && !isSafeNavigationHref(href)) {
        element.removeAttribute('href');
      }

      const safeHref = element.getAttribute('href');
      if (safeHref && !isLocalFragmentHref(safeHref)) {
        element.setAttribute('rel', 'noopener noreferrer');
        element.setAttribute('target', '_blank');
      } else {
        element.removeAttribute('rel');
        element.removeAttribute('target');
      }
    }
  });

  return template.innerHTML;
}

export function htmlToText(html: string): string {
  const element = document.createElement('div');
  element.innerHTML = html;
  return element.textContent ?? '';
}
