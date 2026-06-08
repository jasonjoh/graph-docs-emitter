// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program, getDoc } from '@typespec/compiler';
import { formatTypeName } from '../utils/type-formatter.js';
import { getModelProperties } from '../utils/model-properties.js';
import { escapeMarkdownCell } from '../utils/markdown.js';

export const TODO_DESCRIPTION = '**TODO: Add description**';

export interface PropertiesTableProps {
  program: Program;
  model: Model;
}

/**
 * Renders a Markdown table of model properties (non-navigation only).
 * Navigation properties are handled by RelationshipsTable.
 */
export function PropertiesTable(props: PropertiesTableProps): Children {
  const rows: { name: string; row: string }[] = [];

  for (const [name, property] of getModelProperties(
    props.program,
    props.model,
    'exclude',
  )) {
    const typeName = formatTypeName(property.type);
    const description = escapeMarkdownCell(
      getDoc(props.program, property) ?? TODO_DESCRIPTION,
    );
    rows.push({ name, row: `| ${name} | ${typeName} | ${description} |` });
  }

  if (rows.length === 0) return [];

  rows.sort((a, b) => a.name.localeCompare(b.name));

  return [
    '\n## Properties\n\n',
    '| Property | Type | Description |\n',
    '|:--|:--|:--|\n',
    ...rows.map((r) => r.row + '\n'),
  ];
}

/**
 * Returns true if any property on the model (or its base) lacks a description.
 */
export function hasMissingDescriptions(
  program: Program,
  model: Model,
): boolean {
  for (const [, property] of getModelProperties(program, model)) {
    if (!getDoc(program, property)) return true;
  }
  return false;
}

/**
 * Returns an HTML comment warning if any properties lack descriptions.
 */
export function getTodoComment(program: Program, model: Model): string {
  if (!hasMissingDescriptions(program, model)) return '';
  return (
    '<!-- This file contains placeholder descriptions ("TODO: Add description") because\n' +
    '     the source TypeSpec file is missing documentation comments for some properties.\n' +
    '     Please update the TypeSpec source with the missing descriptions and regenerate. -->\n\n'
  );
}
