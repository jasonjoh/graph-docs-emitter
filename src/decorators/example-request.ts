// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DecoratorContext, Operation, Program } from '@typespec/compiler';

const exampleRequestKey = Symbol('exampleRequest');

/**
 * `@exampleRequest` — attach a custom example request body to an operation.
 */
export function $exampleRequest(
  context: DecoratorContext,
  target: Operation,
  value: unknown,
): void {
  context.program.stateMap(exampleRequestKey).set(target, value);
}

/**
 * Retrieve the custom example request body for an operation, if any.
 */
export function getExampleRequest(
  program: Program,
  target: Operation,
): unknown | undefined {
  return program.stateMap(exampleRequestKey).get(target);
}
