// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

/** @jsxImportSource @alloy-js/core */
import { EmitContext, Model, Namespace } from '@typespec/compiler';
import { Output, SourceFile, SourceDirectory } from '@alloy-js/core';
import { writeOutput } from '@typespec/emitter-framework';
import { GraphDocsEmitterOptions } from './lib.js';
import { collectGraphTypes, GraphRouteInfo } from './utils/type-collector.js';
import {
  resolveOperationsFromRoute,
  ResolvedOperation,
} from './utils/operation-resolver.js';
import { getTypeFilename, getMethodFilename } from './utils/filename.js';
import { ResourceTypePage } from './components/ResourceTypePage.jsx';
import { ApiMethodPage } from './components/ApiMethodPage.jsx';
import { EnumsPage } from './components/EnumsPage.jsx';
import { ComplexTypePage } from './components/ComplexTypePage.jsx';
import {
  getPublicNamespaceName,
  hasPublicNamespace,
} from '@microsoft/typespec-msgraph';

export { $lib } from './lib.js';
export {
  $exampleRequest,
  getExampleRequest,
} from './decorators/example-request.js';
export {
  $exampleResponse,
  getExampleResponse,
} from './decorators/example-response.js';

const DEFAULT_NAMESPACE = 'microsoft.graph';

export async function $onEmit(context: EmitContext<GraphDocsEmitterOptions>) {
  if (context.program.compilerOptions.noEmit) {
    return;
  }

  const program = context.program;
  const types = collectGraphTypes(program);

  // Resolve the public namespace name for display
  function getNamespaceForType(ns: Namespace | undefined): string {
    while (ns) {
      if (hasPublicNamespace(program, ns)) {
        return getPublicNamespaceName(program, ns) ?? DEFAULT_NAMESPACE;
      }
      ns = ns.namespace;
    }
    return DEFAULT_NAMESPACE;
  }

  // Build operation map: entity name -> operations
  const operationsByEntity = new Map<string, ResolvedOperation[]>();
  for (const route of types.routes) {
    const entityName = getEntityNameForRoute(route);
    const ops = resolveOperationsFromRoute(
      program,
      route,
      types.entities,
      entityName,
    );
    if (entityName && ops.length > 0) {
      const existing = operationsByEntity.get(entityName) ?? [];
      existing.push(...ops);
      operationsByEntity.set(entityName, existing);
    }
  }

  // Collect all method pages
  const methodPages: {
    filename: string;
    op: ResolvedOperation;
    ns: string;
    entityModel: Model | undefined;
  }[] = [];
  for (const [entityName, ops] of operationsByEntity) {
    for (const op of ops) {
      const entity = types.entities.find((e) => e.name === entityName);
      const ns = entity
        ? getNamespaceForType(entity.model.namespace)
        : DEFAULT_NAMESPACE;
      methodPages.push({
        filename: getMethodFilename(op, entityName),
        op,
        ns,
        entityModel: entity?.model,
      });
    }
  }

  const outputDir = context.options['output-dir'] ?? context.emitterOutputDir;
  const msDate = context.options['ms-date'];
  const author = context.options.author;
  const apiVersion = context.options['api-version'] ?? 'v1.0';

  await writeOutput(
    program,
    <Output>
      <SourceDirectory path='resources'>
        {types.entities.map((entity) => (
          <SourceFile path={getTypeFilename(entity.name)} filetype='md'>
            <ResourceTypePage
              program={program}
              model={entity.model}
              description={entity.description}
              namespace={getNamespaceForType(entity.model.namespace)}
              operations={operationsByEntity.get(entity.name) ?? []}
              getMethodFilename={(op) => getMethodFilename(op, entity.name)}
              apiVersion={apiVersion}
              msDate={msDate}
              author={author}
            />
          </SourceFile>
        ))}
        {types.complexTypes.map((complex) => (
          <SourceFile path={getTypeFilename(complex.name)} filetype='md'>
            <ComplexTypePage
              program={program}
              model={complex.model}
              description={complex.description}
              namespace={getNamespaceForType(complex.model.namespace)}
              apiVersion={apiVersion}
              msDate={msDate}
              author={author}
            />
          </SourceFile>
        ))}
        {types.enums.length > 0 && (
          <SourceFile path='enums.md' filetype='md'>
            <EnumsPage
              program={program}
              enums={types.enums}
              namespace={DEFAULT_NAMESPACE}
              apiVersion={apiVersion}
              msDate={msDate}
              author={author}
            />
          </SourceFile>
        )}
      </SourceDirectory>
      <SourceDirectory path='api'>
        {methodPages.map((page) => (
          <SourceFile path={page.filename} filetype='md'>
            <ApiMethodPage
              operation={page.op}
              namespace={page.ns}
              filename={page.filename}
              entityModel={page.entityModel}
              program={program}
              apiVersion={apiVersion}
              msDate={msDate}
              author={author}
            />
          </SourceFile>
        ))}
      </SourceDirectory>
    </Output>,
    outputDir,
  );
}

/**
 * Extract the resource entity name from a route interface's source interfaces.
 * Route interfaces extend Resource<T> or Collection<T>, where T is the entity model.
 */
function getEntityNameForRoute(route: GraphRouteInfo): string | undefined {
  for (const src of route.iface.sourceInterfaces) {
    if (src.templateMapper?.args) {
      for (const arg of src.templateMapper.args) {
        if (arg.entityKind === 'Type' && arg.kind === 'Model' && arg.name) {
          return arg.name;
        }
      }
    }
  }
  return undefined;
}
