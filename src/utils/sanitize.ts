import DOMPurify from 'dompurify';

const networkAttributes = ['src', 'srcset', 'href', 'poster', 'data'];

function isNetworkReference(value: string): boolean {
  const clean = value.trim();
  return /^https?:\/\//i.test(clean) || /^\/\//.test(clean);
}

export function sanitizeHtml(html: string): string {
  const sanitized = DOMPurify.sanitize(html, {
    FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'link', 'meta'],
    FORBID_ATTR: ['onerror', 'onload', 'onclick', 'onmouseover', 'style'],
  });

  const template = document.createElement('template');
  template.innerHTML = sanitized;

  template.content.querySelectorAll('*').forEach((element) => {
    networkAttributes.forEach((attribute) => {
      const value = element.getAttribute(attribute);
      if (value && isNetworkReference(value)) {
        element.removeAttribute(attribute);
      }
    });

    if (element.tagName.toLowerCase() === 'a') {
      element.setAttribute('rel', 'noopener noreferrer');
      element.setAttribute('target', '_blank');
    }
  });

  return template.innerHTML;
}

export function htmlToText(html: string): string {
  const element = document.createElement('div');
  element.innerHTML = html;
  return element.textContent ?? '';
}
