// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program } from '@typespec/compiler';
import { isContains } from '@microsoft/typespec-msgraph';
import { formatJsonValue } from '../utils/type-formatter.js';

export interface JsonRepresentationProps {
  program: Program;
  model: Model;
}

/**
 * Renders the ## JSON representation section with a sample JSON object.
 */
export function JsonRepresentation(props: JsonRepresentationProps): Children {
  const lines: string[] = [];
  lines.push('{');

  const entries: string[] = [];

  // Add @odata.type
  if (props.model.name) {
    entries.push(`  "@odata.type": "#microsoft.graph.${props.model.name}"`);
  }

  for (const [name, property] of props.model.properties) {
    // Skip containment nav properties from JSON representation
    if (isContains(props.program, property)) continue;
    const jsonValue = formatJsonValue(property.type);
    entries.push(`  "${name}": ${jsonValue}`);
  }

  lines.push(entries.join(',\n'));
  lines.push('}');

  const jsonBlock = lines.join('\n');

  return [
    '\n## JSON representation\n\n',
    'The following JSON representation shows the resource type.\n\n',
    '<!-- {\n  "blockType": "resource",\n  "@odata.type": "microsoft.graph.',
    props.model.name ?? '',
    '"\n} -->\n',
    '```json\n',
    jsonBlock,
    '\n```\n',
  ];
}
