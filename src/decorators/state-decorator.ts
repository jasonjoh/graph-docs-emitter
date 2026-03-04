// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DecoratorContext, Operation, Program } from '@typespec/compiler';

/**
 * Creates a decorator setter/getter pair backed by a program stateMap.
 * Used to attach arbitrary metadata to TypeSpec operations.
 */
export function createStateDecorator(name: string): {
  set: (context: DecoratorContext, target: Operation, value: unknown) => void;
  get: (program: Program, target: Operation) => unknown | undefined;
} {
  const key = Symbol(name);
  return {
    set(context, target, value) {
      context.program.stateMap(key).set(target, value);
    },
    get(program, target) {
      return program.stateMap(key).get(target);
    },
  };
}
