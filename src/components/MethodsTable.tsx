// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { ResolvedOperation } from '../utils/operation-resolver.js';

export interface MethodsTableProps {
  operations: ResolvedOperation[];
  /** Function to generate the filename for an operation's doc page */
  getMethodFilename: (op: ResolvedOperation) => string;
}

/**
 * Renders the ## Methods section with a table linking to individual method pages.
 */
export function MethodsTable(props: MethodsTableProps): Children {
  if (props.operations.length === 0) return [];

  const rows = props.operations.map((op) => {
    const filename = props.getMethodFilename(op);
    const returnType = op.returnTypeName ?? 'None';
    const description = op.description ?? '';
    return `| [${op.name}](../api/${filename}) | ${returnType} | ${description} |`;
  });

  return [
    '\n## Methods\n\n',
    '| Method | Return Type | Description |\n',
    '|:---|:---|:---|\n',
    ...rows.map((r) => r + '\n'),
  ];
}
