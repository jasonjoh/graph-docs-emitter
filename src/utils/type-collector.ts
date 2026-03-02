// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

import {
  Program,
  Model,
  Enum,
  Interface,
  Namespace,
  getDoc,
  navigateProgram,
} from '@typespec/compiler';
import {
  isEntity,
  isComplex,
  isOperationParameters,
  hasPublicNamespace,
  hasGraphRoute,
  getGraphRoutePaths,
} from '@microsoft/typespec-msgraph';

/**
 * A Graph entity model with its metadata.
 */
export interface GraphEntityInfo {
  name: string;
  model: Model;
  description: string | undefined;
}

/**
 * A Graph complex type with its metadata.
 */
export interface GraphComplexTypeInfo {
  name: string;
  model: Model;
  description: string | undefined;
}

/**
 * A Graph enum with its metadata.
 */
export interface GraphEnumInfo {
  name: string;
  enumType: Enum;
  description: string | undefined;
}

/**
 * A route interface with its graph route path.
 */
export interface GraphRouteInfo {
  path: string;
  iface: Interface;
}

/**
 * Collected Graph types from the TypeSpec program, organized for emission.
 */
export interface CollectedTypes {
  entities: GraphEntityInfo[];
  complexTypes: GraphComplexTypeInfo[];
  enums: GraphEnumInfo[];
  routes: GraphRouteInfo[];
}

/**
 * Check if a namespace (or any of its ancestors) has @publicNamespace.
 */
function isInPublicNamespace(
  program: Program,
  ns: Namespace | undefined,
): boolean {
  while (ns) {
    if (hasPublicNamespace(program, ns)) {
      return true;
    }
    ns = ns.namespace;
  }
  return false;
}

/**
 * Walk the TypeSpec program and collect all Graph-relevant types.
 * Only includes types defined in @publicNamespace namespaces,
 * filtering out library/reference types.
 */
export function collectGraphTypes(program: Program): CollectedTypes {
  const entities: GraphEntityInfo[] = [];
  const complexTypes: GraphComplexTypeInfo[] = [];
  const enums: GraphEnumInfo[] = [];
  const routes: GraphRouteInfo[] = [];

  navigateProgram(program, {
    model(model) {
      if (!model.namespace || !isInPublicNamespace(program, model.namespace)) {
        return;
      }
      // Skip template declarations (only emit instances)
      if (model.node?.kind === undefined) return;
      // Skip operation parameter models — they are inlined
      if (isOperationParameters(program, model)) return;
      // Skip anonymous models
      if (!model.name || model.name === '') return;

      const description = getDoc(program, model) ?? undefined;

      if (isEntity(program, model)) {
        entities.push({ name: model.name, model, description });
      } else if (isComplex(program, model)) {
        complexTypes.push({ name: model.name, model, description });
      }
    },

    enum(enumType) {
      if (
        !enumType.namespace ||
        !isInPublicNamespace(program, enumType.namespace)
      ) {
        return;
      }
      if (!enumType.name || enumType.name === '') return;

      const description = getDoc(program, enumType) ?? undefined;
      enums.push({ name: enumType.name, enumType, description });
    },

    interface(iface) {
      if (!hasGraphRoute(program, iface)) return;
      if (!iface.namespace || !isInPublicNamespace(program, iface.namespace)) {
        return;
      }

      const paths = getGraphRoutePaths(program, iface);
      if (paths && paths.length > 0) {
        routes.push({ path: paths[0], iface });
      }
    },
  });

  return { entities, complexTypes, enums, routes };
}
