// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Program, Operation, Model, getDoc } from '@typespec/compiler';
import {
  GraphRouteInfo,
  GraphEntityInfo,
  GlobalOperationInterfaceInfo,
  normalizeDescription,
} from './type-collector.js';
import {
  getAgsAttributes,
  isGlobalOperation,
} from '@microsoft/typespec-msgraph';

/**
 * Represents a resolved HTTP operation ready for documentation generation.
 */
export interface ResolvedOperation {
  /** Display name for the operation (e.g., "Get user", "List calendars") */
  name: string;
  /** HTTP verb */
  httpMethod: string;
  /** The full graph route path (e.g., "users/{id}") */
  routePath: string;
  /** The resource type name this operation targets */
  resourceTypeName: string;
  /** Brief description from doc comments */
  description: string | undefined;
  /** The operation kind for doc page generation */
  docKind: DocOperationKind;
  /** For actions/functions, the operation name */
  actionOrFunctionName?: string;
  /** The parent entity name (for collection operations) */
  parentEntityName?: string;
  /** The return type name */
  returnTypeName: string | undefined;
  /** Original TypeSpec Operation (for reading decorators) */
  typespecOperation?: Operation;
  /** The request body model for actions/functions (from operation parameters) */
  requestBodyModel?: Model;
}

/**
 * Operation kinds for determining the documentation page type and filename.
 */
export enum DocOperationKind {
  /** GET returning a collection */
  ListCollection = 'list',
  /** GET returning a single resource */
  GetResource = 'get',
  /** POST that creates an item in a collection */
  PostCreate = 'post',
  /** POST action (not creating) */
  Action = 'action',
  /** GET function */
  Function = 'function',
  /** PATCH update */
  Update = 'update',
  /** DELETE */
  Delete = 'delete',
}

/**
 * Known operation name patterns from GraphOps templates and LocalOps.
 * Maps the TypeSpec operation name to its HTTP method and doc kind.
 */
const OPERATION_PATTERNS: Record<
  string,
  { httpMethod: string; docKind: DocOperationKind }
> = {
  // Collection operations
  GetCollection: {
    httpMethod: 'GET',
    docKind: DocOperationKind.ListCollection,
  },
  GetPagedCollection: {
    httpMethod: 'GET',
    docKind: DocOperationKind.ListCollection,
  },
  get: { httpMethod: 'GET', docKind: DocOperationKind.ListCollection },

  // Resource operations
  GetResource: { httpMethod: 'GET', docKind: DocOperationKind.GetResource },

  // Create operations
  Post: { httpMethod: 'POST', docKind: DocOperationKind.PostCreate },
  post: { httpMethod: 'POST', docKind: DocOperationKind.PostCreate },

  // Update operations
  Patch: { httpMethod: 'PATCH', docKind: DocOperationKind.Update },
  patch: { httpMethod: 'PATCH', docKind: DocOperationKind.Update },
  PatchNoResponse: { httpMethod: 'PATCH', docKind: DocOperationKind.Update },
  PatchWithResponse: { httpMethod: 'PATCH', docKind: DocOperationKind.Update },
  Put: { httpMethod: 'PUT', docKind: DocOperationKind.Update },
  PutWithResponse: { httpMethod: 'PUT', docKind: DocOperationKind.Update },

  // Delete operations
  Delete: { httpMethod: 'DELETE', docKind: DocOperationKind.Delete },
  delete: { httpMethod: 'DELETE', docKind: DocOperationKind.Delete },

  // Action operations
  Action: { httpMethod: 'POST', docKind: DocOperationKind.Action },

  // Function operations
  Function: { httpMethod: 'GET', docKind: DocOperationKind.Function },
};

interface ParsedRouteSegments {
  /** Last non-parameter segment (the resource name) */
  resource: string;
  /** Previous non-parameter segment (the parent entity), if any */
  parent: string | undefined;
  /** True if the route targets a collection (no {id} at the end) */
  isCollection: boolean;
}

/**
 * Parse a route path into its resource segment, parent segment,
 * and whether it targets a collection or single resource.
 *
 * E.g., "copilot/conversations/{conversationId}/messages"
 *   -> { resource: "messages", parent: "conversations", isCollection: true }
 */
function parseRouteSegments(routePath: string): ParsedRouteSegments {
  const segments = routePath.split('/');
  const isCollection = !segments[segments.length - 1].startsWith('{');

  let resource = segments[0];
  let parent: string | undefined;

  for (let i = segments.length - 1; i >= 0; i--) {
    if (!segments[i].startsWith('{')) {
      resource = segments[i];
      for (let j = i - 1; j >= 0; j--) {
        if (!segments[j].startsWith('{')) {
          parent = segments[j];
          break;
        }
      }
      break;
    }
  }

  return { resource, parent, isCollection };
}

/**
 * Resolve operations from a route interface into documentation-ready descriptors.
 * Inspects the interface's operations and uses name-based pattern matching
 * against known GraphOps/LocalOps template names.
 */
export function resolveOperationsFromRoute(
  program: Program,
  route: GraphRouteInfo,
  _entities: GraphEntityInfo[],
  entityName: string | undefined,
): ResolvedOperation[] {
  const resolved: ResolvedOperation[] = [];
  const routePath = route.path;
  const { resource: resourceSegment, isCollection } =
    parseRouteSegments(routePath);

  for (const [opName, operation] of route.iface.operations) {
    const resolved_op = resolveOperation(
      program,
      operation,
      opName,
      routePath,
      resourceSegment,
      isCollection,
      entityName,
    );
    if (resolved_op) {
      resolved.push(resolved_op);
    }
  }

  // Also check source interfaces (for "extends Resource<T>" / "extends Collection<T>")
  for (const sourceIface of route.iface.sourceInterfaces) {
    for (const [opName, operation] of sourceIface.operations) {
      const resolved_op = resolveOperation(
        program,
        operation,
        opName,
        routePath,
        resourceSegment,
        isCollection,
        entityName,
      );
      if (resolved_op) {
        resolved.push(resolved_op);
      }
    }
  }

  return resolved;
}

/**
 * Resolve operations from a global operation interface (no @graphRoute)
 * into documentation-ready descriptors. Only includes operations
 * decorated with @globalOperation.
 */
export function resolveGlobalOperations(
  program: Program,
  globalOpInterface: GlobalOperationInterfaceInfo,
  routePath: string,
): ResolvedOperation[] {
  const resolved: ResolvedOperation[] = [];
  const { resource: resourceSegment, isCollection } =
    parseRouteSegments(routePath);
  const entityName = globalOpInterface.entityName;

  for (const [opName, operation] of globalOpInterface.iface.operations) {
    if (!isGlobalOperation(program, operation)) continue;
    const resolved_op = resolveOperation(
      program,
      operation,
      opName,
      routePath,
      resourceSegment,
      isCollection,
      entityName,
    );
    if (resolved_op) {
      resolved.push(resolved_op);
    }
  }

  // Also check source interfaces
  for (const sourceIface of globalOpInterface.iface.sourceInterfaces) {
    for (const [opName, operation] of sourceIface.operations) {
      if (!isGlobalOperation(program, operation)) continue;
      const resolved_op = resolveOperation(
        program,
        operation,
        opName,
        routePath,
        resourceSegment,
        isCollection,
        entityName,
      );
      if (resolved_op) {
        resolved.push(resolved_op);
      }
    }
  }

  return resolved;
}

/**
 * Build a route index mapping entity names to their collection and resource routes.
 */
export function buildEntityRouteIndex(
  routes: GraphRouteInfo[],
  getEntityName: (route: GraphRouteInfo) => string | undefined,
): Map<string, { collectionRoutes: string[]; resourceRoutes: string[] }> {
  const index = new Map<
    string,
    { collectionRoutes: string[]; resourceRoutes: string[] }
  >();

  for (const route of routes) {
    const entityName = getEntityName(route);
    if (!entityName) continue;

    if (!index.has(entityName)) {
      index.set(entityName, { collectionRoutes: [], resourceRoutes: [] });
    }
    const entry = index.get(entityName)!;
    const { isCollection } = parseRouteSegments(route.path);
    if (isCollection) {
      entry.collectionRoutes.push(route.path);
    } else {
      entry.resourceRoutes.push(route.path);
    }
  }

  return index;
}

function resolveOperation(
  program: Program,
  operation: Operation,
  opName: string,
  routePath: string,
  resourceSegment: string,
  isCollection: boolean,
  entityName: string | undefined,
): ResolvedOperation | undefined {
  const description = normalizeDescription(
    getDoc(program, operation) ?? undefined,
  );

  // Skip hidden operations
  const attrs = getAgsAttributes(program, operation);
  if (attrs.get('IsHidden') === 'true') return undefined;

  // Check the operation name against known patterns
  // Prefer source operation name (template name) over the alias name
  let pattern = undefined;
  if (operation.sourceOperation) {
    pattern = OPERATION_PATTERNS[operation.sourceOperation.name];
  }
  if (!pattern) {
    pattern = OPERATION_PATTERNS[opName];
  }

  if (pattern) {
    // For actions and functions, use the operation name and append to route
    if (
      pattern.docKind === DocOperationKind.Action ||
      pattern.docKind === DocOperationKind.Function
    ) {
      const returnTypeName = getActionReturnTypeName(operation);
      const displayName = entityName ? `${entityName}: ${opName}` : opName;
      const requestBodyModel = getActionParametersModel(operation);
      return {
        name: displayName,
        httpMethod: pattern.httpMethod,
        routePath: `${routePath}/${opName}`,
        resourceTypeName: resourceSegment,
        description,
        docKind: pattern.docKind,
        actionOrFunctionName: opName,
        returnTypeName,
        typespecOperation: operation,
        requestBodyModel,
      };
    }

    // Adjust list vs get for "get" operations on collection routes
    let docKind = pattern.docKind;
    if (
      opName === 'get' &&
      isCollection &&
      pattern.docKind === DocOperationKind.ListCollection
    ) {
      docKind = DocOperationKind.ListCollection;
    }

    const { parent: parentSegment } = parseRouteSegments(routePath);
    // Derive return type from entity name and operation kind
    const returnTypeName = getReturnTypeForCrud(docKind, entityName);
    const displayName = getCrudDisplayName(docKind, entityName);

    return {
      name: displayName,
      httpMethod: pattern.httpMethod,
      routePath,
      resourceTypeName: resourceSegment,
      description,
      docKind,
      parentEntityName: parentSegment,
      returnTypeName,
      typespecOperation: operation,
    };
  }

  // If not a known pattern, treat as an action
  const displayName = entityName ? `${entityName}: ${opName}` : opName;
  return {
    name: displayName,
    httpMethod: 'POST',
    routePath: `${routePath}/${opName}`,
    resourceTypeName: resourceSegment,
    description,
    docKind: DocOperationKind.Action,
    actionOrFunctionName: opName,
    returnTypeName: getActionReturnTypeName(operation),
    typespecOperation: operation,
    requestBodyModel: getActionParametersModel(operation),
  };
}

/**
 * Get the display name for a CRUD operation matching the template format.
 * Uses capitalized verb + entity name: "List configurationMonitor"
 */
function getCrudDisplayName(
  docKind: DocOperationKind,
  entityName: string | undefined,
): string {
  const name = entityName ?? 'resource';
  switch (docKind) {
    case DocOperationKind.ListCollection:
      return `List ${name}`;
    case DocOperationKind.GetResource:
      return `Get ${name}`;
    case DocOperationKind.PostCreate:
      return `Create ${name}`;
    case DocOperationKind.Update:
      return `Update ${name}`;
    case DocOperationKind.Delete:
      return `Delete ${name}`;
    default:
      return `${docKind} ${name}`;
  }
}

/**
 * Get the standard CRUD description for an operation when no doc comment exists.
 */
export function getStandardCrudDescription(
  docKind: DocOperationKind,
  entityName: string | undefined,
): string {
  const name = entityName ?? 'resource';
  switch (docKind) {
    case DocOperationKind.ListCollection:
      return `Get a list of ${name} objects.`;
    case DocOperationKind.GetResource:
      return `Retrieve the properties and relationships of a ${name} object.`;
    case DocOperationKind.PostCreate:
      return `Create a new ${name} object.`;
    case DocOperationKind.Update:
      return `Update a ${name} object.`;
    case DocOperationKind.Delete:
      return `Delete a ${name} object.`;
    default:
      return `${docKind} ${name}.`;
  }
}

/**
 * Derive the return type name for CRUD operations from the entity name
 * and operation kind. This avoids parsing complex Union response types.
 */
function getReturnTypeForCrud(
  docKind: DocOperationKind,
  entityName: string | undefined,
): string | undefined {
  if (!entityName) return undefined;

  switch (docKind) {
    case DocOperationKind.ListCollection:
      return `${entityName} collection`;
    case DocOperationKind.GetResource:
    case DocOperationKind.PostCreate:
    case DocOperationKind.Update:
      return entityName;
    case DocOperationKind.Delete:
      return undefined;
    default:
      return entityName;
  }
}

/**
 * Extract the return type name for action/function operations.
 * Inspects the operation's source template arguments (e.g., Action<TReturnType>)
 * to find the actual return type.
 */
function getActionReturnTypeName(operation: Operation): string | undefined {
  // Check the source operation's template mapper for the return type argument
  if (operation.sourceOperation?.templateMapper?.args) {
    for (const arg of operation.sourceOperation.templateMapper.args) {
      if (arg.entityKind === 'Type') {
        if (arg.kind === 'Intrinsic' && arg.name === 'void') {
          return undefined;
        }
        if (arg.kind === 'Model' && arg.name) {
          return arg.name;
        }
        if (arg.kind === 'Scalar' && arg.name) {
          return arg.name;
        }
      }
    }
  }

  // Fallback: check the operation's direct return type
  const returnType = operation.returnType;
  if (returnType.kind === 'Model') {
    return returnType.name || undefined;
  }
  if (returnType.kind === 'Intrinsic' && returnType.name === 'void') {
    return undefined;
  }
  return undefined;
}

/**
 * Extract the action/function parameters model from an operation.
 * Uses the operation's parameters model which contains the
 * request body properties for actions.
 */
function getActionParametersModel(operation: Operation): Model | undefined {
  const params = operation.parameters;
  if (params && params.properties.size > 0) {
    return params;
  }
  return undefined;
}
