import type { ReactNode } from 'react';

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function countMatches(text: string, query: string): number {
  if (!query.trim()) {
    return 0;
  }

  return text.match(new RegExp(escapeRegExp(query.trim()), 'gi'))?.length ?? 0;
}

export function highlightText(text: string, query: string): ReactNode {
  const cleanQuery = query.trim();

  if (!cleanQuery) {
    return text;
  }

  const pattern = new RegExp(`(${escapeRegExp(cleanQuery)})`, 'gi');
  const parts = text.split(pattern);

  return parts.map((part, index) => {
    if (part.toLowerCase() === cleanQuery.toLowerCase()) {
      return (
        <mark className="ldv-highlight" key={`${part}-${index}`}>
          {part}
        </mark>
      );
    }

    return part;
  });
}
