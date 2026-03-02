// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { YamlFrontMatter } from './YamlFrontMatter.jsx';
import { ResolvedOperation } from '../utils/operation-resolver.js';

export interface ApiMethodPageProps {
  operation: ResolvedOperation;
  namespace: string;
  msDate?: string;
  author?: string;
}

/**
 * Renders a full API method page matching the learn.microsoft.com format.
 */
export function ApiMethodPage(props: ApiMethodPageProps): Children {
  const op = props.operation;
  const title = op.name;
  const desc = op.description ?? `${op.name}.`;

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
    `Namespace: ${props.namespace}\n\n`,
    desc + '\n',
    '\n## HTTP request\n\n',
    '<!-- { "blockType": "ignored" } -->\n',
    '```http\n',
    `${op.httpMethod.toUpperCase()} /${op.routePath}\n`,
    '```\n',
    '\n## Request headers\n\n',
    '| Name | Description |\n',
    '|:---|:---|\n',
    '| Authorization | Bearer {token}. Required. |\n',
    '| Content-Type | application/json. Required for methods with a request body. |\n',
    renderRequestBody(op),
    renderResponse(op),
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
      'If successful, this method returns a `204 No Content` response code.\n',
    ];
  }

  // Handle collection types (e.g., "appQuotaSettings collection")
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
