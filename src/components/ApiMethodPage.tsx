// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Model, Program, getDoc } from '@typespec/compiler';
import {
  isContains,
  isComputed,
  isImmutable,
  isReadOnly,
} from '@microsoft/typespec-msgraph';
import { DEFAULT_NAMESPACE } from '../utils/graph-metadata.js';
import { YamlFrontMatter } from './YamlFrontMatter.jsx';
import {
  ResolvedOperation,
  DocOperationKind,
  getStandardCrudDescription,
} from '../utils/operation-resolver.js';
import { formatTypeName, formatJsonValue } from '../utils/type-formatter.js';
import { escapeMarkdownCell } from '../utils/markdown.js';
import { getExampleRequest } from '../decorators/example-request.js';
import { getExampleResponse } from '../decorators/example-response.js';

export interface ApiMethodPageProps {
  operation: ResolvedOperation;
  namespace: string;
  filename: string;
  program: Program;
  entityModel?: Model;
  apiVersion?: string;
  msDate?: string;
  author?: string;
}

/**
 * Get the success HTTP status code for an operation kind.
 */
function getSuccessStatusCode(op: ResolvedOperation): string {
  if (op.docKind === DocOperationKind.Delete || !op.returnTypeName) {
    return '204 No Content';
  }
  if (op.docKind === DocOperationKind.PostCreate) {
    return '201 Created';
  }
  return '200 OK';
}

/**
 * Renders a full API method page matching the learn.microsoft.com format.
 */
export function ApiMethodPage(props: ApiMethodPageProps): Children {
  const op = props.operation;
  const title = op.name;
  const desc = getDescription(op);
  const isBeta = props.apiVersion === 'beta';
  const permissionName = props.filename.replace(/\.md$/, '') + '-permissions';

  return [
    <YamlFrontMatter
      title={title}
      description={desc}
      docType='apiPageType'
      msDate={props.msDate}
      author={props.author}
    />,
    '\n',
    `# ${title}\n\n`,
    '<!-- markdownlint-disable MD024 -->\n\n',
    isBeta
      ? '[!INCLUDE [beta-disclaimer](../../includes/beta-disclaimer.md)]\n\n'
      : '',
    desc + '\n',
    '\n## Permissions\n\n',
    'Choose the permission or permissions marked as least privileged for this API. Use a higher privileged permission or permissions [only if your app requires it](/graph/permissions-overview#best-practices-for-using-microsoft-graph-permissions). For details about delegated and application permissions, see [Permission types](/graph/permissions-overview#permission-types). To learn more about these permissions, see the [permissions reference](/graph/permissions-reference).\n\n',
    `<!-- {\n  "blockType": "permissions",\n  "name": "${permissionName}"\n}\n-->\n\n`,
    `[!INCLUDE [permissions-table](../includes/permissions/${permissionName}.md)]\n`,
    '\n## HTTP request\n\n',
    '```http\n',
    `${op.httpMethod.toUpperCase()} /${op.routePath}\n`,
    '```\n',
    renderQueryParameters(op),
    '\n## Request headers\n\n',
    '| Name | Description |\n',
    '|:--|:--|\n',
    '| Authorization | `Bearer {token}.` Required. Learn more about [authentication and authorization](/graph/auth/auth-concepts). |\n',
    renderContentTypeHeader(op),
    renderRequestBody(op, props.program),
    renderResponse(op),
    renderExample(op, props),
  ];
}

function getDescription(op: ResolvedOperation): string {
  if (op.description) return op.description;
  // Use standard CRUD descriptions as fallback
  if (
    op.docKind !== DocOperationKind.Action &&
    op.docKind !== DocOperationKind.Function
  ) {
    // Extract entity name from the display name (e.g., "Get configurationMonitor")
    const entityFromName = op.name.split(' ').slice(1).join(' ') || undefined;
    const entityName = op.returnTypeName ?? entityFromName;
    return getStandardCrudDescription(op.docKind, entityName);
  }
  return `${op.name}.`;
}

function renderContentTypeHeader(op: ResolvedOperation): Children {
  const method = op.httpMethod.toUpperCase();
  if (method === 'GET' || method === 'DELETE') {
    return [];
  }
  return ['| Content-Type | `application/json`. Required. |\n'];
}

function renderQueryParameters(op: ResolvedOperation): Children {
  if (op.httpMethod.toUpperCase() !== 'GET') return [];
  return [
    '\n## Optional query parameters\n\n',
    'This method supports [OData query parameters](/graph/query-parameters) to help customize the response.\n',
  ];
}

function renderRequestBody(op: ResolvedOperation, program: Program): Children {
  if (
    op.httpMethod.toUpperCase() === 'GET' ||
    op.httpMethod.toUpperCase() === 'DELETE'
  ) {
    return [
      '\n## Request body\n\n',
      "Don't supply a request body for this method.\n",
    ];
  }

  // For actions with a request body model, show the parameters table
  if (op.requestBodyModel && op.requestBodyModel.properties.size > 0) {
    const rows: string[] = [];
    for (const [name, property] of op.requestBodyModel.properties) {
      const typeName = formatTypeName(property.type);
      const description = escapeMarkdownCell(getDoc(program, property) ?? '');
      rows.push(`| ${name} | ${typeName} | ${description} |`);
    }

    if (rows.length > 0) {
      return [
        '\n## Request body\n\n',
        'In the request body, supply a JSON representation of the parameters.\n\n',
        '| Property | Type | Description |\n',
        '|:--|:--|:--|\n',
        ...rows.map((r) => r + '\n'),
      ];
    }
  }

  return [
    '\n## Request body\n\n',
    'In the request body, supply a JSON representation of the resource.\n',
  ];
}

function renderResponse(op: ResolvedOperation): Children {
  const returnType = op.returnTypeName;
  if (!returnType) {
    return [
      '\n## Response\n\n',
      `If successful, this method returns a \`204 No Content\` response code. It doesn't return anything in the response body.\n`,
    ];
  }

  const isCollection = returnType.endsWith(' collection');
  const baseType = isCollection
    ? returnType.replace(' collection', '')
    : returnType;
  const linkPath = `../resources/${baseType.toLowerCase()}.md`;
  const statusCode = getSuccessStatusCode(op);

  return [
    '\n## Response\n\n',
    `If successful, this method returns a \`${statusCode}\` response code and a [${returnType}](${linkPath}) object in the response body.\n`,
  ];
}

function renderExample(
  op: ResolvedOperation,
  props: ApiMethodPageProps,
): Children {
  const returnType = op.returnTypeName;
  const statusCode = getSuccessStatusCode(op);
  const requestName = props.filename.replace(/\.md$/, '');
  const ns = props.namespace;
  const hasRequestBody =
    op.httpMethod.toUpperCase() !== 'GET' &&
    op.httpMethod.toUpperCase() !== 'DELETE';

  // Check for custom examples from decorators
  const customRequest = op.typespecOperation
    ? getExampleRequest(props.program, op.typespecOperation)
    : undefined;
  const customResponse = op.typespecOperation
    ? getExampleResponse(props.program, op.typespecOperation)
    : undefined;

  const requestBody = hasRequestBody
    ? getRequestBodyJson(props, customRequest)
    : '';
  const responseBody = getResponseBodyJson(props, op, ns, customResponse);

  const result: (string | Children)[] = [
    '\n## Example\n\n',
    '### Request\n\n',
    'The following example shows a request.\n\n',
    `<!-- {\n  "blockType": "request",\n  "name": "${requestName}"\n}\n-->\n\n`,
    '```http\n',
    `${op.httpMethod.toUpperCase()} https://graph.microsoft.com/${props.apiVersion ?? 'v1.0'}/${op.routePath}\n`,
    hasRequestBody ? 'Content-type: application/json\n' : '',
    requestBody ? '\n' + requestBody + '\n' : '',
    '```\n',
    '\n### Response\n\n',
    'The following example shows the response.',
    returnType
      ? ' The response shown here might be shortened for readability.'
      : '',
    '\n\n',
  ];

  if (returnType) {
    const isCollection = op.docKind === DocOperationKind.ListCollection;
    const baseType = returnType.replace(' collection', '');
    const odataType = `${ns}.${baseType}`;
    const collectionProp = isCollection ? ',\n  "isCollection": true' : '';
    result.push(
      `<!-- {\n  "blockType": "response",\n  "truncated": true,\n  "@odata.type": "${odataType}"${collectionProp}\n}\n-->\n\n`,
    );
    result.push('```http\n');
    result.push(`HTTP/1.1 ${statusCode}\n`);
    result.push('Content-type: application/json\n');
    if (responseBody) {
      result.push('\n' + responseBody + '\n');
    }
    result.push('```\n');
  } else {
    result.push(
      `<!-- {\n  "blockType": "response",\n  "truncated": true\n}\n-->\n\n`,
    );
    result.push('```http\n');
    result.push(`HTTP/1.1 ${statusCode}\n`);
    result.push('```\n');
  }

  return result;
}

/**
 * Get request body JSON — custom example or auto-generated.
 */
function getRequestBodyJson(
  props: ApiMethodPageProps,
  customExample: unknown | undefined,
): string {
  if (customExample) {
    return JSON.stringify(customExample, null, 2);
  }
  // For actions with a request body model, use the action parameters
  if (props.operation.requestBodyModel) {
    return buildJsonBodyFromModel(
      props.operation.requestBodyModel,
      props.namespace,
    );
  }
  if (props.entityModel) {
    return buildJsonBody(props, true);
  }
  return '';
}

/**
 * Get response body JSON — custom example or auto-generated.
 */
function getResponseBodyJson(
  props: ApiMethodPageProps,
  op: ResolvedOperation,
  ns: string,
  customExample: unknown | undefined,
): string {
  if (customExample) {
    return JSON.stringify(customExample, null, 2);
  }
  if (!op.returnTypeName || !props.entityModel) return '';
  if (op.docKind === DocOperationKind.ListCollection) {
    return buildCollectionJsonBody(props, ns);
  }
  return buildJsonBody(props);
}

/**
 * Build a JSON body from the entity model's properties.
 * When forRequest is true, excludes computed, immutable, and
 * readOnly properties (not settable by the caller).
 */
function buildJsonBody(props: ApiMethodPageProps, forRequest = false): string {
  const model = props.entityModel;
  if (!model) return '{}';

  const ns = props.namespace ?? DEFAULT_NAMESPACE;
  const program = props.program;
  const entries: string[] = [];

  if (model.name) {
    entries.push(`  "@odata.type": "#${ns}.${model.name}"`);
  }

  for (const [name, property] of model.properties) {
    if (isContains(program, property)) continue;
    if (
      forRequest &&
      (isComputed(program, property) ||
        isImmutable(program, property) ||
        isReadOnly(program, property))
    ) {
      continue;
    }
    entries.push(`  "${name}": ${formatJsonValue(property.type)}`);
  }

  return '{\n' + entries.join(',\n') + '\n}';
}

/**
 * Build a collection response JSON body.
 */
function buildCollectionJsonBody(
  props: ApiMethodPageProps,
  ns: string,
): string {
  const model = props.entityModel;
  if (!model) return '{}';

  const innerBody = buildJsonBody(props)
    .split('\n')
    .map((l) => '    ' + l)
    .join('\n');

  return (
    '{\n' +
    `  "@odata.context": "https://graph.microsoft.com/$metadata#${ns}",\n` +
    '  "value": [\n' +
    innerBody +
    '\n  ]\n' +
    '}'
  );
}

/**
 * Build a JSON body from an arbitrary model (e.g., action parameters).
 */
function buildJsonBodyFromModel(model: Model, _ns: string): string {
  const entries: string[] = [];

  for (const [name, property] of model.properties) {
    entries.push(`  "${name}": ${formatJsonValue(property.type)}`);
  }

  if (entries.length === 0) return '{}';
  return '{\n' + entries.join(',\n') + '\n}';
}
