// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore jasonjoh testitem testentity testwidget

/** @jsxImportSource @alloy-js/core */
import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { renderTree, printTree } from '@alloy-js/core';
import { createGraphDocsTestRunner } from '../test-host.js';
import { collectGraphTypes } from '../../src/utils/type-collector.js';
import {
  resolveOperationsFromRoute,
  DocOperationKind,
} from '../../src/utils/operation-resolver.js';
import { ApiMethodPage } from '../../src/components/ApiMethodPage.jsx';

function renderToString(jsx: unknown): string {
  const tree = renderTree(jsx as Parameters<typeof renderTree>[0]);
  return printTree(tree);
}

let runner: BasicTestRunner;

beforeEach(async () => {
  runner = await createGraphDocsTestRunner();
});

/**
 * Helper: compile a TSP spec with an entity + routes and resolve operations.
 */
async function compileAndResolve(tsp: string, entityName?: string) {
  await runner.compile(tsp);
  const types = collectGraphTypes(runner.program);
  const ops = types.routes.flatMap((route) =>
    resolveOperationsFromRoute(
      runner.program,
      route,
      types.entities,
      entityName,
    ),
  );
  return { types, ops };
}

const STANDARD_SPEC = `
  using MsGraph;

  @publicNamespace("microsoft.graph")
  namespace microsoft.graph {
    /** A test entity representing a widget. */
    @entity model testWidget {
      /** The unique identifier. */
      @readOnly @computed @key id: string;
      /** The display name. */
      @requiredForCreate displayName: string;
      /** The current count. */
      @requiredForCreate count: int32;
      /** The creator. */
      @computed createdBy: string;
    }

    @graphRoute("widgets")
    interface widgetCollection extends Collection<testWidget> {
      @select @top @skip getAll is GraphOps.GetPagedCollection;
      post is GraphOps.Post;
    }

    @graphRoute("widgets/{id}")
    interface widgetById extends Resource<testWidget> {
      get is GraphOps.GetResource;
      patch is GraphOps.PatchNoResponse;
      delete is GraphOps.Delete;
    }
  }
`;

describe('ApiMethodPage', () => {
  describe('GET single resource', () => {
    it('renders correct page structure', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
          apiVersion='v1.0'
          msDate='01/15/2025'
          author='jasonjoh'
        />,
      );

      expect(result).toContain('doc_type: apiPageType');
      expect(result).toContain('author: jasonjoh');
      expect(result).toContain('## Permissions');
      expect(result).toContain('## HTTP request');
      expect(result).toContain('GET /widgets/{id}');
      expect(result).toContain('## Request headers');
      expect(result).toContain('## Request body');
      expect(result).toContain('## Response');
      expect(result).toContain('## Example');
    });

    it('renders optional query parameters section', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('## Optional query parameters');
      expect(result).toContain('OData query parameters');
    });

    it('renders "don\'t supply request body" for GET', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain("Don't supply a request body");
    });

    it('does not render Content-Type header for GET', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).not.toContain('Content-Type | `application/json`');
    });

    it('renders 200 OK response', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('`200 OK`');
      expect(result).toContain('HTTP/1.1 200 OK');
    });

    it('renders response example with all properties', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      // Response should contain all properties
      expect(result).toContain('"id": "String"');
      expect(result).toContain('"displayName": "String"');
      expect(result).toContain('"count": 0');
      expect(result).toContain('"createdBy": "String"');
    });

    it('does not include request body in GET example', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      // The request block should not have Content-type for GET
      const requestSection = result
        .split('### Request')[1]
        .split('### Response')[0];
      expect(requestSection).not.toContain('Content-type: application/json');
    });
  });

  describe('POST create', () => {
    it('renders Content-Type header for POST', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const postOp = ops.find(
        (o) => o.docKind === DocOperationKind.PostCreate,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={postOp}
          namespace='microsoft.graph'
          filename='widget-post-widgets.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('Content-Type | `application/json`');
    });

    it('renders "supply JSON" for POST request body', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const postOp = ops.find(
        (o) => o.docKind === DocOperationKind.PostCreate,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={postOp}
          namespace='microsoft.graph'
          filename='widget-post-widgets.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('supply a JSON representation');
    });

    it('renders 201 Created response', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const postOp = ops.find(
        (o) => o.docKind === DocOperationKind.PostCreate,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={postOp}
          namespace='microsoft.graph'
          filename='widget-post-widgets.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('`201 Created`');
      expect(result).toContain('HTTP/1.1 201 Created');
    });

    it('excludes computed/readOnly properties from request body example', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const postOp = ops.find(
        (o) => o.docKind === DocOperationKind.PostCreate,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={postOp}
          namespace='microsoft.graph'
          filename='widget-post-widgets.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      // Request body should exclude @computed properties (id, createdBy)
      const requestSection = result
        .split('### Request')[1]
        .split('### Response')[0];
      expect(requestSection).not.toContain('"id"');
      expect(requestSection).not.toContain('"createdBy"');
      // But should include non-computed properties
      expect(requestSection).toContain('"displayName"');
      expect(requestSection).toContain('"count"');
    });

    it('does not render query parameters section for POST', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const postOp = ops.find(
        (o) => o.docKind === DocOperationKind.PostCreate,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={postOp}
          namespace='microsoft.graph'
          filename='widget-post-widgets.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).not.toContain('## Optional query parameters');
    });
  });

  describe('PATCH update', () => {
    it('renders Content-Type header for PATCH', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const patchOp = ops.find((o) => o.docKind === DocOperationKind.Update)!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={patchOp}
          namespace='microsoft.graph'
          filename='testwidget-update.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('Content-Type | `application/json`');
    });

    it('renders PATCH in HTTP request section', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const patchOp = ops.find((o) => o.docKind === DocOperationKind.Update)!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={patchOp}
          namespace='microsoft.graph'
          filename='testwidget-update.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('PATCH /widgets/{id}');
    });
  });

  describe('DELETE', () => {
    it('renders 204 No Content response with no body', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const deleteOp = ops.find((o) => o.docKind === DocOperationKind.Delete)!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={deleteOp}
          namespace='microsoft.graph'
          filename='testwidget-delete.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('`204 No Content`');
      expect(result).toContain("doesn't return anything in the response body");
    });

    it('renders "don\'t supply request body" for DELETE', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const deleteOp = ops.find((o) => o.docKind === DocOperationKind.Delete)!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={deleteOp}
          namespace='microsoft.graph'
          filename='testwidget-delete.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain("Don't supply a request body");
    });

    it('does not render Content-Type header for DELETE', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const deleteOp = ops.find((o) => o.docKind === DocOperationKind.Delete)!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={deleteOp}
          namespace='microsoft.graph'
          filename='testwidget-delete.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).not.toContain('Content-Type | `application/json`');
    });

    it('renders simple response example with no body', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const deleteOp = ops.find((o) => o.docKind === DocOperationKind.Delete)!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={deleteOp}
          namespace='microsoft.graph'
          filename='testwidget-delete.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('HTTP/1.1 204 No Content');
      // Response should NOT contain Content-type or JSON body
      const responseSection = result.split('### Response')[1];
      expect(responseSection).not.toContain('Content-type: application/json');
    });
  });

  describe('GET collection (list)', () => {
    it('renders collection response with value wrapper', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const listOp = ops.find(
        (o) => o.docKind === DocOperationKind.ListCollection,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={listOp}
          namespace='microsoft.graph'
          filename='widget-list-widgets.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('"@odata.context"');
      expect(result).toContain('"value": [');
    });

    it('renders isCollection in response HTML comment', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const listOp = ops.find(
        (o) => o.docKind === DocOperationKind.ListCollection,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={listOp}
          namespace='microsoft.graph'
          filename='widget-list-widgets.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('"isCollection": true');
    });
  });

  describe('permissions section', () => {
    it('renders permissions block comment and INCLUDE directive', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('"blockType": "permissions"');
      expect(result).toContain('"name": "testwidget-get-permissions"');
      expect(result).toContain(
        '[!INCLUDE [permissions-table](../includes/permissions/testwidget-get-permissions.md)]',
      );
    });
  });

  describe('example section HTML comments', () => {
    it('renders request block comment with filename-based name', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('"blockType": "request"');
      expect(result).toContain('"name": "testwidget-get"');
    });

    it('renders response block comment with @odata.type', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('"blockType": "response"');
      expect(result).toContain('"@odata.type": "microsoft.graph.testWidget"');
    });

    it('renders full URL with api-version in request example', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
          apiVersion='beta'
        />,
      );

      expect(result).toContain(
        'GET https://graph.microsoft.com/beta/widgets/{id}',
      );
    });

    it('defaults api-version to v1.0 in request example', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain(
        'GET https://graph.microsoft.com/v1.0/widgets/{id}',
      );
    });
  });

  describe('beta disclaimer', () => {
    it('renders beta disclaimer when apiVersion is beta', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
          apiVersion='beta'
        />,
      );

      expect(result).toContain('beta-disclaimer');
    });

    it('omits beta disclaimer when apiVersion is v1.0', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
          apiVersion='v1.0'
        />,
      );

      expect(result).not.toContain('beta-disclaimer');
    });
  });

  describe('action operations', () => {
    it('renders action page with POST and request body', async () => {
      const { types, ops } = await compileAndResolve(
        `
        using MsGraph;

        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            /** The unique identifier. */
            @readOnly @computed @key id: string;
            /** The display name. */
            @computed displayName: string;
          }

          @operationParameters model testActionParams {
            value: string;
          }

          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            doSomething is GraphOps.Action<TActionParams=testActionParams, TReturnType=testItem>;
          }
        }
      `,
        'testItem',
      );

      const actionOp = ops.find((o) => o.docKind === DocOperationKind.Action)!;
      const entity = types.entities.find((e) => e.name === 'testItem')!;

      const result = renderToString(
        <ApiMethodPage
          operation={actionOp}
          namespace='microsoft.graph'
          filename='testitem-dosomething.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('POST /items/{id}');
      expect(result).toContain('Content-Type | `application/json`');
      expect(result).toContain('supply a JSON representation');
    });
  });

  describe('response type link', () => {
    it('links to resource page for single return type', async () => {
      const { types, ops } = await compileAndResolve(
        STANDARD_SPEC,
        'testWidget',
      );
      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      const entity = types.entities.find((e) => e.name === 'testWidget')!;

      const result = renderToString(
        <ApiMethodPage
          operation={getOp}
          namespace='microsoft.graph'
          filename='testwidget-get.md'
          program={runner.program}
          entityModel={entity.model}
        />,
      );

      expect(result).toContain('[testWidget](../resources/testwidget.md)');
    });
  });
});
