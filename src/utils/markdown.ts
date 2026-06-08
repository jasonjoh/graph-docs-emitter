// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Escape a string for use inside a Markdown table cell.
 * Replaces pipe characters and newlines that would break table formatting.
 */
export function escapeMarkdownCell(text: string): string {
  return text.replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>');
}
