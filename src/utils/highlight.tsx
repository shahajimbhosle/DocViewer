import type { ReactNode } from 'react';

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export interface TextMatchRange {
  start: number;
  end: number;
}

export function findMatchRanges(text: string, query: string): TextMatchRange[] {
  const cleanQuery = query.trim();

  if (!cleanQuery) {
    return [];
  }

  const pattern = new RegExp(escapeRegExp(cleanQuery), 'gi');
  const ranges: TextMatchRange[] = [];
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(text))) {
    ranges.push({
      start: match.index,
      end: match.index + match[0].length,
    });

    if (match[0].length === 0) {
      pattern.lastIndex += 1;
    }
  }

  return ranges;
}

export function countMatches(text: string, query: string): number {
  return findMatchRanges(text, query).length;
}

export function highlightText(text: string, query: string): ReactNode {
  const ranges = findMatchRanges(text, query);

  if (ranges.length === 0) {
    return text;
  }

  const parts: ReactNode[] = [];
  let cursor = 0;

  ranges.forEach((range, index) => {
    if (range.start > cursor) {
      parts.push(text.slice(cursor, range.start));
    }

    parts.push(
      <mark className="ldv-highlight" key={`${range.start}-${range.end}-${index}`}>
        {text.slice(range.start, range.end)}
      </mark>,
    );
    cursor = range.end;
  });

  if (cursor < text.length) {
    parts.push(text.slice(cursor));
  }

  return parts;
}
