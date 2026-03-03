// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program } from '@typespec/compiler';
import { YamlFrontMatter } from './YamlFrontMatter.jsx';
import { PropertiesTable } from './PropertiesTable.jsx';
import { RelationshipsTable } from './RelationshipsTable.jsx';
import { JsonRepresentation } from './JsonRepresentation.jsx';
import { MethodsTable } from './MethodsTable.jsx';
import { ResolvedOperation } from '../utils/operation-resolver.js';

export interface ResourceTypePageProps {
  program: Program;
  model: Model;
  description: string | undefined;
  namespace: string;
  operations: ResolvedOperation[];
  getMethodFilename: (op: ResolvedOperation) => string;
  apiVersion?: string;
  msDate?: string;
  author?: string;
}

/**
 * Renders a full resource type page matching the learn.microsoft.com format.
 */
export function ResourceTypePage(props: ResourceTypePageProps): Children {
  const title = `${props.model.name} resource type`;
  const desc = props.description ?? `Represents a ${props.model.name}.`;
  const isBeta = props.apiVersion === 'beta';

  return [
    <YamlFrontMatter
      title={title}
      description={desc}
      docType='resourcePageType'
      msDate={props.msDate}
      author={props.author}
    />,
    '\n',
    `# ${title}\n\n`,
    `Namespace: ${props.namespace}\n\n`,
    isBeta
      ? '[!INCLUDE [beta-disclaimer](../../includes/beta-disclaimer.md)]\n\n'
      : '',
    desc + '\n',
    <MethodsTable
      operations={props.operations}
      getMethodFilename={props.getMethodFilename}
    />,
    <PropertiesTable program={props.program} model={props.model} />,
    <RelationshipsTable program={props.program} model={props.model} />,
    <JsonRepresentation
      program={props.program}
      model={props.model}
      namespace={props.namespace}
    />,
  ];
}
