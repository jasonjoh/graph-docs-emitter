// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { YamlFrontMatter } from './YamlFrontMatter.jsx';
import {
  ResolvedOperation,
  DocOperationKind,
  getStandardCrudDescription,
} from '../utils/operation-resolver.js';

export interface ApiMethodPageProps {
  operation: ResolvedOperation;
  namespace: string;
  apiVersion?: string;
  msDate?: string;
  author?: string;
}

/**
 * Renders a full API method page matching the learn.microsoft.com format.
 */
export function ApiMethodPage(props: ApiMethodPageProps): Children {
  const op = props.operation;
  const title = op.name;
  const desc = getDescription(op);
  const isBeta = props.apiVersion === 'beta';

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
    'Choose the permission or permissions marked as least privileged for this API. Use a higher privileged permission or permissions [only if your app requires it](/graph/permissions-overview#best-practices-for-using-microsoft-graph-permissions). For details about delegated and application permissions, see [Permission types](/graph/permissions-overview#permission-types). To learn more about these permissions, see the [permissions reference](/graph/permissions-reference).\n',
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
    renderRequestBody(op),
    renderResponse(op),
    renderExample(op),
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

function renderRequestBody(op: ResolvedOperation): Children {
  if (
    op.httpMethod.toUpperCase() === 'GET' ||
    op.httpMethod.toUpperCase() === 'DELETE'
  ) {
    return [
      '\n## Request body\n\n',
      "Don't supply a request body for this method.\n",
    ];
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
  const statusCode = op.docKind === 'post' ? '201 Created' : '200 OK';

  return [
    '\n## Response\n\n',
    `If successful, this method returns a \`${statusCode}\` response code and a [${returnType}](${linkPath}) object in the response body.\n`,
  ];
}

function renderExample(op: ResolvedOperation): Children {
  const returnType = op.returnTypeName;
  const statusCode =
    op.docKind === 'delete'
      ? '204 No Content'
      : op.docKind === 'post'
        ? '201 Created'
        : '200 OK';

  return [
    '\n## Example\n\n',
    '### Request\n\n',
    'The following example shows a request.\n\n',
    '```http\n',
    `${op.httpMethod.toUpperCase()} /${op.routePath}\n`,
    '```\n',
    '\n### Response\n\n',
    'The following example shows the response.',
    returnType
      ? ' The response shown here might be shortened for readability.'
      : '',
    '\n\n',
    '```http\n',
    `HTTP/1.1 ${statusCode}\n`,
    returnType ? 'Content-type: application/json\n' : '',
    '```\n',
  ];
}
