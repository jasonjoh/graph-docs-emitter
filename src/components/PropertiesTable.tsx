// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program, getDoc } from '@typespec/compiler';
import { isContains } from '@microsoft/typespec-msgraph';
import { formatTypeName } from '../utils/type-formatter.js';

export interface PropertiesTableProps {
  program: Program;
  model: Model;
}

/**
 * Renders a Markdown table of model properties (non-navigation only).
 * Navigation properties are handled by RelationshipsTable.
 */
export function PropertiesTable(props: PropertiesTableProps): Children {
  const rows: string[] = [];

  for (const [name, property] of props.model.properties) {
    // Skip navigation/containment properties — they go in Relationships
    if (isContains(props.program, property)) continue;

    const typeName = formatTypeName(property.type);
    const description = getDoc(props.program, property) ?? '';
    rows.push(`| ${name} | ${typeName} | ${description} |`);
  }

  // Also include inherited properties from base model
  if (props.model.baseModel) {
    for (const [name, property] of props.model.baseModel.properties) {
      if (isContains(props.program, property)) continue;
      const typeName = formatTypeName(property.type);
      const description = getDoc(props.program, property) ?? '';
      rows.push(`| ${name} | ${typeName} | ${description} |`);
    }
  }

  if (rows.length === 0) return [];

  rows.sort((a, b) => a.localeCompare(b));

  return [
    '\n## Properties\n\n',
    '| Property | Type | Description |\n',
    '|:--|:--|:--|\n',
    ...rows.map((r) => r + '\n'),
  ];
}
