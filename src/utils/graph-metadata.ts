// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

import { Program, ModelProperty, Type, getDoc } from '@typespec/compiler';
import {
  isEntity,
  isComplex,
  isContains,
  isOwnerless,
  isComputed,
  isReadOnly,
  isRequiredForCreate,
  isImmutable,
  isOperationParameters,
  hasPublicNamespace,
  getPublicNamespaceName,
  getGraphRoutePaths,
  hasGraphRoute,
  getAddressUrl,
} from '@microsoft/typespec-msgraph';

// Re-export decorator accessors for convenient use across the emitter.
export {
  isEntity,
  isComplex,
  isContains,
  isOwnerless,
  isComputed,
  isReadOnly,
  isRequiredForCreate,
  isImmutable,
  isOperationParameters,
  hasPublicNamespace,
  getPublicNamespaceName,
  getGraphRoutePaths,
  hasGraphRoute,
  getAddressUrl,
};

/**
 * Get the documentation string for a type, if available.
 */
export function getDescription(
  program: Program,
  type: Type,
): string | undefined {
  return getDoc(program, type) ?? undefined;
}

/**
 * Classify a model property for documentation purposes.
 */
export interface PropertyClassification {
  isComputed: boolean;
  isReadOnly: boolean;
  isImmutable: boolean;
  isRequiredForCreate: boolean;
  isContainment: boolean;
  isNullable: boolean;
}

/**
 * Get the full classification of a model property, combining multiple
 * decorator checks into a single object.
 */
export function classifyProperty(
  program: Program,
  property: ModelProperty,
): PropertyClassification {
  return {
    isComputed: isComputed(program, property),
    isReadOnly: isReadOnly(program, property),
    isImmutable: isImmutable(program, property),
    isRequiredForCreate: isRequiredForCreate(program, property),
    isContainment: isContains(program, property),
    isNullable: isPropertyNullable(property),
  };
}

/**
 * Check if a property type is nullable (union with null).
 */
function isPropertyNullable(property: ModelProperty): boolean {
  const type = property.type;
  if (type.kind === 'Union') {
    for (const variant of type.variants.values()) {
      if (variant.type.kind === 'Intrinsic' && variant.type.name === 'null') {
        return true;
      }
    }
  }
  return false;
}
