// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  createTypeSpecLibrary,
  JSONSchemaType,
  paramMessage,
} from '@typespec/compiler';

export interface GraphDocsEmitterOptions {
  'api-version'?: string;
  'output-dir'?: string;
  'ms-date'?: string;
  author?: string;
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
        'The Microsoft Graph API version (e.g., v1.0, beta). Used for display purposes in generated documentation.',
    },
    'output-dir': {
      type: 'string',
      nullable: true,
      description:
        'Override the output directory for generated files. Defaults to the compiler emitter output directory.',
    },
    'ms-date': {
      type: 'string',
      nullable: true,
      description:
        'Override the ms.date value in YAML front matter (MM/DD/YYYY format). Defaults to today.',
    },
    author: {
      type: 'string',
      nullable: true,
      description:
        "The author's GitHub username. Set as the `author` field in YAML front matter.",
    },
  },
  required: [],
};

export const $lib = createTypeSpecLibrary({
  name: '@microsoft/typespec-graph-docs-emitter',
  diagnostics: {
    'unknown-operation-pattern': {
      severity: 'warning',
      messages: {
        default: paramMessage`Operation "${'opName'}" does not match any known GraphOps pattern and will be skipped. Use a recognized template (e.g., GraphOps.Action, GraphOps.GetResource) or add it to the emitter's OPERATION_PATTERNS.`,
      },
    },
  },
  emitter: {
    options: EmitterOptionsSchema,
  },
});
