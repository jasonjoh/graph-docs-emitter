// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import {
  ResolvedOperation,
  DocOperationKind,
} from '../utils/operation-resolver.js';
import { escapeMarkdownCell } from '../utils/markdown.js';

export interface MethodsTableProps {
  operations: ResolvedOperation[];
  /** Function to generate the filename for an operation's doc page */
  getMethodFilename: (op: ResolvedOperation) => string;
}

/** CRUD sort order: List, Create, Get, Update, Delete first */
const CRUD_ORDER: Record<string, number> = {
  [DocOperationKind.ListCollection]: 0,
  [DocOperationKind.PostCreate]: 1,
  [DocOperationKind.GetResource]: 2,
  [DocOperationKind.Update]: 3,
  [DocOperationKind.Delete]: 4,
};

/**
 * Renders the ## Methods section with a table linking to individual method pages.
 */
export function MethodsTable(props: MethodsTableProps): Children {
  if (props.operations.length === 0) return [];

  // Sort: CRUD operations first in standard order, then others alphabetically
  const sorted = [...props.operations].sort((a, b) => {
    const aOrder = CRUD_ORDER[a.docKind] ?? 99;
    const bOrder = CRUD_ORDER[b.docKind] ?? 99;
    if (aOrder !== bOrder) return aOrder - bOrder;
    return a.name.localeCompare(b.name);
  });

  const rows = sorted.map((op) => {
    const filename = props.getMethodFilename(op);
    const returnType = op.returnTypeName ? `\`${op.returnTypeName}\`` : 'None';
    const description = escapeMarkdownCell(op.description ?? '');
    return `| [${op.name}](../api/${filename}) | ${returnType} | ${description} |`;
  });

  return [
    '\n## Methods\n\n',
    '| Method | Return Type | Description |\n',
    '|:--|:--|:--|\n',
    ...rows.map((r) => r + '\n'),
  ];
}
