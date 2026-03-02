// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Program, Operation, getDoc } from '@typespec/compiler';
import { GraphRouteInfo, GraphEntityInfo } from './type-collector.js';

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
): ResolvedOperation | undefined {
  const description = getDoc(program, operation) ?? undefined;

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
    const returnTypeName = getReturnTypeName(operation);

    return {
      name: `${docKind} ${resourceSegment}`,
      httpMethod: pattern.httpMethod,
      routePath,
      resourceTypeName: resourceSegment,
      description,
      docKind,
      parentEntityName: parentSegment,
      returnTypeName,
    };
  }

  // If not a known pattern, treat as an action
  return {
    name: opName,
    httpMethod: 'POST',
    routePath,
    resourceTypeName: resourceSegment,
    description,
    docKind: DocOperationKind.Action,
    actionOrFunctionName: opName,
    returnTypeName: getReturnTypeName(operation),
  };
}

/**
 * Extract the return type name from an operation.
 */
function getReturnTypeName(operation: Operation): string | undefined {
  const returnType = operation.returnType;
  if (returnType.kind === 'Model') {
    return returnType.name || undefined;
  }
  if (returnType.kind === 'Intrinsic' && returnType.name === 'void') {
    return undefined;
  }
  return returnType.kind;
}
