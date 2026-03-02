// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { EmitContext, Namespace } from '@typespec/compiler';
import { Output, SourceFile, SourceDirectory } from '@alloy-js/core';
import { writeOutput } from '@typespec/emitter-framework';
import { GraphDocsEmitterOptions } from './lib.js';
import { collectGraphTypes } from './utils/type-collector.js';
import {
  resolveOperationsFromRoute,
  ResolvedOperation,
} from './utils/operation-resolver.js';
import { getTypeFilename, getMethodFilename } from './utils/filename.js';
import { ResourceTypePage } from './components/ResourceTypePage.jsx';
import { ApiMethodPage } from './components/ApiMethodPage.jsx';
import { EnumTypePage } from './components/EnumTypePage.jsx';
import { ComplexTypePage } from './components/ComplexTypePage.jsx';
import {
  getPublicNamespaceName,
  hasPublicNamespace,
} from '@microsoft/typespec-msgraph';

export { $lib } from './lib.js';

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
    const ops = resolveOperationsFromRoute(program, route, types.entities);
    for (const op of ops) {
      // Try to find which entity this route targets
      const entityName = findEntityForRoute(
        route.path,
        types.entities.map((e) => e.name),
      );
      if (entityName) {
        const existing = operationsByEntity.get(entityName) ?? [];
        existing.push(op);
        operationsByEntity.set(entityName, existing);
      }
    }
  }

  // Collect all method pages
  const methodPages: { filename: string; op: ResolvedOperation; ns: string }[] =
    [];
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
      });
    }
  }

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
            />
          </SourceFile>
        ))}
        {types.enums.map((enumInfo) => (
          <SourceFile path={getTypeFilename(enumInfo.name)} filetype='md'>
            <EnumTypePage
              program={program}
              enumType={enumInfo.enumType}
              description={enumInfo.description}
              namespace={getNamespaceForType(enumInfo.enumType.namespace)}
            />
          </SourceFile>
        ))}
      </SourceDirectory>
      <SourceDirectory path='api'>
        {methodPages.map((page) => (
          <SourceFile path={page.filename} filetype='md'>
            <ApiMethodPage operation={page.op} namespace={page.ns} />
          </SourceFile>
        ))}
      </SourceDirectory>
    </Output>,
    context.emitterOutputDir,
  );
}

/**
 * Find which entity a route targets by matching the route path against entity names.
 * Uses a simple heuristic: find the entity whose name (lowercased) appears as a
 * segment in the route path.
 */
function findEntityForRoute(
  routePath: string,
  entityNames: string[],
): string | undefined {
  const pathLower = routePath.toLowerCase();
  // Try to find the most specific match (longest entity name in the path)
  let bestMatch: string | undefined;
  let bestLength = 0;

  for (const name of entityNames) {
    const nameLower = name.toLowerCase();
    if (pathLower.includes(nameLower) && name.length > bestLength) {
      bestMatch = name;
      bestLength = name.length;
    }
  }
  return bestMatch;
}
