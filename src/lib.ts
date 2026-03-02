// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createTypeSpecLibrary, JSONSchemaType } from '@typespec/compiler';

export interface GraphDocsEmitterOptions {
  'api-version'?: string;
}

const EmitterOptionsSchema: JSONSchemaType<GraphDocsEmitterOptions> = {
  type: 'object',
  additionalProperties: false,
  properties: {
    'api-version': {
      type: 'string',
      nullable: true,
      default: 'v1.0',
      description:
        'The API version to use in generated documentation (e.g., v1.0, beta).',
    },
  },
  required: [],
};

export const $lib = createTypeSpecLibrary({
  name: '@microsoft/typespec-graph-docs-emitter',
  diagnostics: {},
  emitter: {
    options: EmitterOptionsSchema,
  },
});
