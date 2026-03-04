// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Program, Operation, getDoc } from '@typespec/compiler';
import {
  GraphRouteInfo,
  GraphEntityInfo,
  normalizeDescription,
} from './type-collector.js';
import { getAgsAttributes } from '@microsoft/typespec-msgraph';

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

/**
 * Determine if a route path targets a collection (no {id} at the end)
 * or a single resource (has {id} at the end).
 */
function isCollectionRoute(routePath: string): boolean {
  const segments = routePath.split('/');
  const last = segments[segments.length - 1];
  return !last.startsWith('{');
}

/**
 * Extract the resource name from the last non-parameter segment of a route.
 * E.g., "copilot/conversations/{conversationId}" -> "conversations"
 */
function getResourceSegment(routePath: string): string {
  const segments = routePath.split('/');
  // Walk backwards to find the last non-parameter segment
  for (let i = segments.length - 1; i >= 0; i--) {
    if (!segments[i].startsWith('{')) {
      return segments[i];
    }
  }
  return segments[0];
}

/**
 * Get the parent entity segment from a route path.
 * E.g., "copilot/conversations/{conversationId}/messages" -> "conversations"
 */
function getParentSegment(routePath: string): string | undefined {
  const segments = routePath.split('/');
  // Find the resource segment, then look for the entity segment before it
  for (let i = segments.length - 1; i >= 0; i--) {
    if (!segments[i].startsWith('{')) {
      // This is the resource segment, look for the previous non-param segment
      for (let j = i - 1; j >= 0; j--) {
        if (!segments[j].startsWith('{')) {
          return segments[j];
        }
      }
      return undefined;
    }
  }
  return undefined;
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
  const resourceSegment = getResourceSegment(routePath);
  const isCollection = isCollectionRoute(routePath);

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

    const parentSegment = getParentSegment(routePath);
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
