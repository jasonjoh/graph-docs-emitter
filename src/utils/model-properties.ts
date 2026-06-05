// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Model, ModelProperty, Program } from '@typespec/compiler';
import { isContains } from '@microsoft/typespec-msgraph';

/**
 * Collects model properties (own + inherited from baseModel),
 * filtered by containment relationship.
 *
 * @param containment - 'exclude' returns only non-navigation properties,
 *   'only' returns only @contains navigation properties,
 *   'all' returns everything.
 */
export function getModelProperties(
  program: Program,
  model: Model,
  containment: 'exclude' | 'only' | 'all' = 'all',
): [string, ModelProperty][] {
  const results: [string, ModelProperty][] = [];
  const seen = new Set<string>();

  function collect(properties: Map<string, ModelProperty>) {
    for (const [name, property] of properties) {
      // Skip properties already seen from a more-derived model
      if (seen.has(name)) continue;
      seen.add(name);

      const isNav = isContains(program, property);
      if (containment === 'exclude' && isNav) continue;
      if (containment === 'only' && !isNav) continue;
      results.push([name, property]);
    }
  }

  // Walk the full inheritance chain, own properties first
  let current: Model | undefined = model;
  while (current) {
    collect(current.properties);
    current = current.baseModel;
  }

  return results;
}
