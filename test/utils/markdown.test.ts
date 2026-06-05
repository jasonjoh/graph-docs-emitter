// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, it, expect } from 'vitest';
import { escapeMarkdownCell } from '../../src/utils/markdown.js';

describe('escapeMarkdownCell', () => {
  it('escapes pipe characters', () => {
    expect(escapeMarkdownCell('a | b')).toBe('a \\| b');
  });

  it('replaces newlines with <br>', () => {
    expect(escapeMarkdownCell('line1\nline2')).toBe('line1<br>line2');
  });

  it('handles CRLF line endings', () => {
    expect(escapeMarkdownCell('line1\r\nline2')).toBe('line1<br>line2');
  });

  it('handles multiple pipes and newlines', () => {
    expect(escapeMarkdownCell('a | b\nc | d')).toBe('a \\| b<br>c \\| d');
  });

  it('returns unchanged text when no special characters', () => {
    expect(escapeMarkdownCell('simple text')).toBe('simple text');
  });

  it('handles empty string', () => {
    expect(escapeMarkdownCell('')).toBe('');
  });
});
