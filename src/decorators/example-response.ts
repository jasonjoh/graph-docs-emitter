// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DecoratorContext, Operation, Program } from '@typespec/compiler';

const exampleResponseKey = Symbol('exampleResponse');

/**
 * `@exampleResponse` — attach a custom example response body to an operation.
 */
export function $exampleResponse(
  context: DecoratorContext,
  target: Operation,
  value: unknown,
): void {
  context.program.stateMap(exampleResponseKey).set(target, value);
}

/**
 * Retrieve the custom example response body for an operation, if any.
 */
export function getExampleResponse(
  program: Program,
  target: Operation,
): unknown | undefined {
  return program.stateMap(exampleResponseKey).get(target);
}
