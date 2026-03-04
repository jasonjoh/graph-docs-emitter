// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program } from '@typespec/compiler';
import { formatJsonValue } from '../utils/type-formatter.js';
import { getModelProperties } from '../utils/model-properties.js';
import { DEFAULT_NAMESPACE } from '../utils/graph-metadata.js';

export interface JsonRepresentationProps {
  program: Program;
  model: Model;
  namespace?: string;
}

/**
 * Renders the ## JSON representation section with a sample JSON object.
 */
export function JsonRepresentation(props: JsonRepresentationProps): Children {
  const lines: string[] = [];
  lines.push('{');

  const ns = props.namespace ?? DEFAULT_NAMESPACE;
  const entries: string[] = [];

  // Add @odata.type
  if (props.model.name) {
    entries.push(`  "@odata.type": "#${ns}.${props.model.name}"`);
  }

  for (const [name, property] of getModelProperties(
    props.program,
    props.model,
    'exclude',
  )) {
    const jsonValue = formatJsonValue(property.type);
    entries.push(`  "${name}": ${jsonValue}`);
  }

  lines.push(entries.join(',\n'));
  lines.push('}');

  const jsonBlock = lines.join('\n');
  const odataType = `${ns}.${props.model.name ?? ''}`;

  return [
    '\n## JSON representation\n\n',
    'The following JSON representation shows the resource type.\n\n',
    '<!-- {\n',
    '  "blockType": "resource",\n',
    '  "keyProperty": "id",\n',
    '  "optionalProperties": [],\n',
    `  "@odata.type": "${odataType}"\n`,
    '} -->\n\n',
    '```json\n',
    jsonBlock,
    '\n```\n',
    '\n<!-- {\n',
    '  "type": "#page.annotation",\n',
    `  "description": "${props.model.name ?? ''} resource",\n`,
    '  "keywords": "",\n',
    '  "section": "documentation",\n',
    '  "tocPath": "",\n',
    '  "suppressions": []\n',
    '} -->\n',
  ];
}
