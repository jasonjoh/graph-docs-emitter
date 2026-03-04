// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createStateDecorator } from './state-decorator.js';

const { set: setExampleRequest, get: getExampleRequest } =
  createStateDecorator('exampleRequest');

/**
 * `@exampleRequest` — attach a custom example request body to an operation.
 */
export const $exampleRequest = setExampleRequest;

/**
 * Retrieve the custom example request body for an operation, if any.
 */
export { getExampleRequest };
