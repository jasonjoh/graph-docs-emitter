// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program } from '@typespec/compiler';
import { YamlFrontMatter } from './YamlFrontMatter.jsx';
import { PropertiesTable } from './PropertiesTable.jsx';
import { JsonRepresentation } from './JsonRepresentation.jsx';

export interface ComplexTypePageProps {
  program: Program;
  model: Model;
  description: string | undefined;
  namespace: string;
}

/**
 * Renders a full complex type page.
 * Similar to resource type but without Methods or Relationships sections.
 */
export function ComplexTypePage(props: ComplexTypePageProps): Children {
  const title = `${props.model.name} resource type`;
  const desc = props.description ?? `Represents a ${props.model.name}.`;

  return [
    <YamlFrontMatter
      title={title}
      description={desc}
      docType='resourcePageType'
    />,
    '\n',
    `# ${title}\n\n`,
    `Namespace: ${props.namespace}\n\n`,
    desc + '\n',
    <PropertiesTable program={props.program} model={props.model} />,
    <JsonRepresentation program={props.program} model={props.model} />,
  ];
}
