// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program, getDoc } from '@typespec/compiler';
import { isContains } from '@microsoft/typespec-msgraph';
import { formatTypeName } from '../utils/type-formatter.js';

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
  const rows: string[] = [];

  for (const [name, property] of props.model.properties) {
    // Skip navigation/containment properties — they go in Relationships
    if (isContains(props.program, property)) continue;

    const typeName = formatTypeName(property.type);
    const description = getDoc(props.program, property) ?? TODO_DESCRIPTION;
    rows.push(`| ${name} | ${typeName} | ${description} |`);
  }

  // Also include inherited properties from base model
  if (props.model.baseModel) {
    for (const [name, property] of props.model.baseModel.properties) {
      if (isContains(props.program, property)) continue;
      const typeName = formatTypeName(property.type);
      const description = getDoc(props.program, property) ?? TODO_DESCRIPTION;
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

/**
 * Returns true if any property on the model (or its base) lacks a description.
 */
export function hasMissingDescriptions(
  program: Program,
  model: Model,
): boolean {
  for (const [, property] of model.properties) {
    if (!getDoc(program, property)) return true;
  }
  if (model.baseModel) {
    for (const [, property] of model.baseModel.properties) {
      if (!getDoc(program, property)) return true;
    }
  }
  return false;
}
