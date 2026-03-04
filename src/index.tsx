// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

/** @jsxImportSource @alloy-js/core */
import { EmitContext, Model } from '@typespec/compiler';
import { Output, SourceFile, SourceDirectory } from '@alloy-js/core';
import { writeOutput } from '@typespec/emitter-framework';
import { GraphDocsEmitterOptions } from './lib.js';
import { collectGraphTypes, GraphRouteInfo } from './utils/type-collector.js';
import {
  resolveOperationsFromRoute,
  ResolvedOperation,
} from './utils/operation-resolver.js';
import { getTypeFilename, getMethodFilename } from './utils/filename.js';
import {
  DEFAULT_NAMESPACE,
  getNamespaceForType,
} from './utils/graph-metadata.js';
import { ResourceTypePage } from './components/ResourceTypePage.jsx';
import { ApiMethodPage } from './components/ApiMethodPage.jsx';
import { EnumsPage } from './components/EnumsPage.jsx';
import { ComplexTypePage } from './components/ComplexTypePage.jsx';

export { $lib } from './lib.js';
export {
  $exampleRequest,
  getExampleRequest,
} from './decorators/example-request.js';
export {
  $exampleResponse,
  getExampleResponse,
} from './decorators/example-response.js';

export async function $onEmit(context: EmitContext<GraphDocsEmitterOptions>) {
  if (context.program.compilerOptions.noEmit) {
    return;
  }

  const program = context.program;
  const types = collectGraphTypes(program);

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
  const entityMap = new Map(types.entities.map((e) => [e.name, e]));
  const methodPages: {
    filename: string;
    op: ResolvedOperation;
    ns: string;
    entityModel: Model | undefined;
  }[] = [];
  for (const [entityName, ops] of operationsByEntity) {
    const entity = entityMap.get(entityName);
    const ns = entity
      ? getNamespaceForType(program, entity.model.namespace)
      : DEFAULT_NAMESPACE;
    for (const op of ops) {
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
      <SourceDirectory path={`${apiVersion}/resources`}>
        {types.entities.map((entity) => (
          <SourceFile path={getTypeFilename(entity.name)} filetype='md'>
            <ResourceTypePage
              program={program}
              model={entity.model}
              description={entity.description}
              namespace={getNamespaceForType(program, entity.model.namespace)}
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
              namespace={getNamespaceForType(program, complex.model.namespace)}
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
      <SourceDirectory path={`${apiVersion}/api`}>
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
