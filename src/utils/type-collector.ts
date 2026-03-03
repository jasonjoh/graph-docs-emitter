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
  getSourceLocation,
  navigateProgram,
} from '@typespec/compiler';
import {
  isEntity,
  isComplex,
  isOperationParameters,
  hasPublicNamespace,
  hasGraphRoute,
  getGraphRoutePaths,
  getAgsAttributes,
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
 * Check if a namespace chain starts with 'reference.' — these are
 * reference package types that should not be emitted.
 */
function isReferenceType(ns: Namespace | undefined): boolean {
  // Walk to the root namespace to build the full path
  const parts: string[] = [];
  let current = ns;
  while (current && current.name) {
    parts.unshift(current.name);
    current = current.namespace;
  }
  return parts[0] === 'reference';
}

/**
 * Check if a type is defined in node_modules (library type, not user-defined).
 */
function isLibraryType(type: Model | Enum | Interface): boolean {
  try {
    const loc = getSourceLocation(type);
    if (loc?.file?.path) {
      return loc.file.path.includes('node_modules');
    }
  } catch {
    // If we can't get the source location, assume it's not a library type
  }
  return false;
}

/**
 * Check if a type has `@agsAttribute("IsHidden", "true")`.
 */
function isHidden(program: Program, type: Model | Enum | Interface): boolean {
  const attrs = getAgsAttributes(program, type);
  return attrs.get('IsHidden') === 'true';
}

/**
 * Normalize description text to ensure proper spacing between sentences.
 * TypeSpec's getDoc() may concatenate multiple doc comment blocks
 * without spaces (e.g., "turn.Represents" instead of "turn. Represents").
 */
export function normalizeDescription(
  text: string | undefined,
): string | undefined {
  if (!text) return text;
  // Insert a space after sentence-ending punctuation (.!?) that is
  // immediately followed by an uppercase letter (start of new sentence)
  return text.replace(/([.!?])([A-Z])/g, '$1 $2');
}

/**
 * Walk the TypeSpec program and collect all Graph-relevant types.
 * Only includes types defined in @publicNamespace namespaces,
 * filtering out library/reference types and hidden types.
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
      // Skip reference package types
      if (isReferenceType(model.namespace)) return;
      // Skip library-defined shared models (e.g., entity, dictionary)
      if (isLibraryType(model)) return;
      // Skip template declarations (only emit instances)
      if (model.node?.kind === undefined) return;
      // Skip operation parameter models — they are inlined
      if (isOperationParameters(program, model)) return;
      // Skip anonymous models
      if (!model.name || model.name === '') return;
      // Skip hidden models
      if (isHidden(program, model)) return;

      const description = normalizeDescription(
        getDoc(program, model) ?? undefined,
      );

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
      // Skip reference package types
      if (isReferenceType(enumType.namespace)) return;
      // Skip library-defined shared enums
      if (isLibraryType(enumType)) return;
      if (!enumType.name || enumType.name === '') return;
      // Skip hidden enums
      if (isHidden(program, enumType)) return;

      const description = normalizeDescription(
        getDoc(program, enumType) ?? undefined,
      );
      enums.push({ name: enumType.name, enumType, description });
    },

    interface(iface) {
      if (!hasGraphRoute(program, iface)) return;
      if (!iface.namespace || !isInPublicNamespace(program, iface.namespace)) {
        return;
      }
      // Skip hidden interfaces
      if (isHidden(program, iface)) return;

      const paths = getGraphRoutePaths(program, iface);
      if (paths && paths.length > 0) {
        routes.push({ path: paths[0], iface });
      }
    },
  });

  return { entities, complexTypes, enums, routes };
}
