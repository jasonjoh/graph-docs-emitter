// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program } from '@typespec/compiler';
import { YamlFrontMatter } from './YamlFrontMatter.jsx';
import { PropertiesTable } from './PropertiesTable.jsx';
import { getTodoComment } from './PropertiesTable.jsx';
import { JsonRepresentation } from './JsonRepresentation.jsx';

export interface ComplexTypePageProps {
  program: Program;
  model: Model;
  description: string | undefined;
  namespace: string;
  apiVersion?: string;
  msDate?: string;
  author?: string;
}

/**
 * Renders a full complex type page.
 * Similar to resource type but without Methods or Relationships sections.
 */
export function ComplexTypePage(props: ComplexTypePageProps): Children {
  const title = `${props.model.name} complex type`;
  const desc = props.description ?? `Represents a ${props.model.name}.`;
  const isBeta = props.apiVersion === 'beta';
  const todoComment = getTodoComment(props.program, props.model);

  return [
    <YamlFrontMatter
      title={title}
      description={desc}
      docType='resourcePageType'
      msDate={props.msDate}
      author={props.author}
    />,
    '\n',
    todoComment,
    `# ${title}\n\n`,
    `Namespace: ${props.namespace}\n\n`,
    isBeta
      ? '[!INCLUDE [beta-disclaimer](../../includes/beta-disclaimer.md)]\n\n'
      : '',
    desc + '\n',
    <PropertiesTable program={props.program} model={props.model} />,
    <JsonRepresentation
      program={props.program}
      model={props.model}
      namespace={props.namespace}
    />,
  ];
}
