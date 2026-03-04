// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createStateDecorator } from './state-decorator.js';

const { set: setExampleResponse, get: getExampleResponse } =
  createStateDecorator('exampleResponse');

/**
 * `@exampleResponse` — attach a custom example response body to an operation.
 */
export const $exampleResponse = setExampleResponse;

/**
 * Retrieve the custom example response body for an operation, if any.
 */
export { getExampleResponse };
