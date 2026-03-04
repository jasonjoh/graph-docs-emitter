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

  function collect(properties: Map<string, ModelProperty>) {
    for (const [name, property] of properties) {
      const isNav = isContains(program, property);
      if (containment === 'exclude' && isNav) continue;
      if (containment === 'only' && !isNav) continue;
      results.push([name, property]);
    }
  }

  collect(model.properties);
  if (model.baseModel) {
    collect(model.baseModel.properties);
  }

  return results;
}
