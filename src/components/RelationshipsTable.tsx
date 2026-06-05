// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program, getDoc } from '@typespec/compiler';
import { formatTypeName } from '../utils/type-formatter.js';
import { getModelProperties } from '../utils/model-properties.js';
import { escapeMarkdownCell } from '../utils/markdown.js';
import { TODO_DESCRIPTION } from './PropertiesTable.jsx';

export interface RelationshipsTableProps {
  program: Program;
  model: Model;
}

/**
 * Renders a Markdown table for @contains navigation properties (Relationships).
 */
export function RelationshipsTable(props: RelationshipsTableProps): Children {
  const rows: { name: string; row: string }[] = [];

  for (const [name, property] of getModelProperties(
    props.program,
    props.model,
    'only',
  )) {
    const typeName = formatTypeName(property.type);
    const description = escapeMarkdownCell(
      getDoc(props.program, property) ?? TODO_DESCRIPTION,
    );
    rows.push({ name, row: `| ${name} | ${typeName} | ${description} |` });
  }

  if (rows.length === 0) {
    return ['\n## Relationships\n\n', 'None.\n'];
  }

  rows.sort((a, b) => a.name.localeCompare(b.name));

  return [
    '\n## Relationships\n\n',
    '| Relationship | Type | Description |\n',
    '|:--|:--|:--|\n',
    ...rows.map((r) => r.row + '\n'),
  ];
}
