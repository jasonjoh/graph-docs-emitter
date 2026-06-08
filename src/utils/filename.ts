// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ResolvedOperation, DocOperationKind } from './operation-resolver.js';

/**
 * Remove filesystem-unsafe characters from a name.
 */
function sanitizeFilename(name: string): string {
  return name.replace(/[/\\:*?"<>|\s]/g, '');
}

/**
 * Generate a filename for a resource, enum, or complex type page.
 * Rule: all-lowercase, no spaces or hyphens.
 * E.g., "copilotConversation" -> "copilotconversation.md"
 */
export function getTypeFilename(typeName: string): string {
  return `${sanitizeFilename(typeName).toLowerCase()}.md`;
}

/**
 * Generate a filename for an API method page based on the learn.microsoft.com naming rules.
 *
 * - GET collection:    {parent-entity}-list-{plural-resource-name}.md
 * - GET single:        {resource-name}-get.md
 * - POST (creates):    {parent-entity}-post-{plural-resource-name}.md
 * - POST (action):     {resource-name}-{action-or-function}.md
 * - PATCH:             {resource-name}-update.md
 * - DELETE:            {resource-name}-delete.md
 */
export function getMethodFilename(
  op: ResolvedOperation,
  entityName: string,
): string {
  const resource = sanitizeFilename(entityName).toLowerCase();

  switch (op.docKind) {
    case DocOperationKind.ListCollection: {
      const parent = sanitizeFilename(
        op.parentEntityName ?? entityName,
      ).toLowerCase();
      const collectionName = sanitizeFilename(
        op.resourceTypeName,
      ).toLowerCase();
      return `${parent}-list-${collectionName}.md`;
    }

    case DocOperationKind.GetResource:
      return `${resource}-get.md`;

    case DocOperationKind.PostCreate: {
      const parent = sanitizeFilename(
        op.parentEntityName ?? entityName,
      ).toLowerCase();
      const collectionName = sanitizeFilename(
        op.resourceTypeName,
      ).toLowerCase();
      return `${parent}-post-${collectionName}.md`;
    }

    case DocOperationKind.Action:
    case DocOperationKind.Function: {
      const actionName = sanitizeFilename(
        op.actionOrFunctionName ?? op.name,
      ).toLowerCase();
      return `${resource}-${actionName}.md`;
    }

    case DocOperationKind.Update:
      return `${resource}-update.md`;

    case DocOperationKind.Delete:
      return `${resource}-delete.md`;

    default:
      return `${resource}-${sanitizeFilename(op.name).toLowerCase()}.md`;
  }
}
