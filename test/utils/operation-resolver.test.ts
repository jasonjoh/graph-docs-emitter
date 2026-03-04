// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { createGraphDocsTestRunner } from '../test-host.js';
import { collectGraphTypes } from '../../src/utils/type-collector.js';
import {
  resolveOperationsFromRoute,
  DocOperationKind,
  getStandardCrudDescription,
} from '../../src/utils/operation-resolver.js';

let runner: BasicTestRunner;

beforeEach(async () => {
  runner = await createGraphDocsTestRunner();
});

describe('getStandardCrudDescription', () => {
  it('returns list description', () => {
    expect(
      getStandardCrudDescription(DocOperationKind.ListCollection, 'widget'),
    ).toBe('Get a list of widget objects.');
  });

  it('returns get description', () => {
    expect(
      getStandardCrudDescription(DocOperationKind.GetResource, 'widget'),
    ).toBe('Retrieve the properties and relationships of a widget object.');
  });

  it('returns post description', () => {
    expect(
      getStandardCrudDescription(DocOperationKind.PostCreate, 'widget'),
    ).toBe('Create a new widget object.');
  });

  it('returns update description', () => {
    expect(getStandardCrudDescription(DocOperationKind.Update, 'widget')).toBe(
      'Update a widget object.',
    );
  });

  it('returns delete description', () => {
    expect(getStandardCrudDescription(DocOperationKind.Delete, 'widget')).toBe(
      'Delete a widget object.',
    );
  });

  it('falls back to "resource" when no entity name', () => {
    expect(
      getStandardCrudDescription(DocOperationKind.GetResource, undefined),
    ).toBe('Retrieve the properties and relationships of a resource object.');
  });
});

describe('resolveOperationsFromRoute', () => {
  describe('GetPagedCollection', () => {
    it('resolves GetPagedCollection as ListCollection', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items")
          interface testItems extends Collection<testItem> {
            @select @top @skip getAll is GraphOps.GetPagedCollection;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path === 'items')!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const listOp = ops.find(
        (o) => o.docKind === DocOperationKind.ListCollection,
      );
      expect(listOp).toBeDefined();
      expect(listOp!.httpMethod).toBe('GET');
      expect(listOp!.returnTypeName).toBe('testItem collection');
    });
  });

  describe('Function operations', () => {
    it('resolves Function as GET with function kind', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
            @computed displayName: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            compute is GraphOps.Function<TReturnType=testItem>;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const funcOp = ops.find((o) => o.docKind === DocOperationKind.Function);
      expect(funcOp).toBeDefined();
      expect(funcOp!.httpMethod).toBe('GET');
      expect(funcOp!.actionOrFunctionName).toBe('compute');
      expect(funcOp!.routePath).toContain('/compute');
    });
  });

  describe('return types', () => {
    it('returns entity name for GetResource', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      expect(getOp.returnTypeName).toBe('testItem');
    });

    it('returns undefined for Delete', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            delete is GraphOps.Delete;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const deleteOp = ops.find((o) => o.docKind === DocOperationKind.Delete)!;
      expect(deleteOp.returnTypeName).toBeUndefined();
    });

    it('returns collection type for ListCollection', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items")
          interface testItems extends Collection<testItem> {
            @select getAll is GraphOps.GetPagedCollection;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path === 'items')!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const listOp = ops.find(
        (o) => o.docKind === DocOperationKind.ListCollection,
      )!;
      expect(listOp.returnTypeName).toBe('testItem collection');
    });

    it('returns undefined for all types when entityName not provided', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        undefined,
      );

      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      expect(getOp.returnTypeName).toBeUndefined();
    });
  });

  describe('display names', () => {
    it('generates CRUD display names from entity name', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
            patch is GraphOps.PatchNoResponse;
            delete is GraphOps.Delete;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      expect(
        ops.find((o) => o.docKind === DocOperationKind.GetResource)!.name,
      ).toBe('Get testItem');
      expect(ops.find((o) => o.docKind === DocOperationKind.Update)!.name).toBe(
        'Update testItem',
      );
      expect(ops.find((o) => o.docKind === DocOperationKind.Delete)!.name).toBe(
        'Delete testItem',
      );
    });

    it('uses "resource" fallback when no entity name', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        undefined,
      );

      expect(
        ops.find((o) => o.docKind === DocOperationKind.GetResource)!.name,
      ).toBe('Get resource');
    });

    it('generates action display name with entity prefix', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @operationParameters model testActionParams {
            value: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            doSomething is GraphOps.Action<TActionParams=testActionParams, TReturnType=testItem>;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const actionOp = ops.find((o) => o.docKind === DocOperationKind.Action)!;
      expect(actionOp.name).toBe('testItem: doSomething');
    });
  });

  describe('typespecOperation field', () => {
    it('populates typespecOperation for CRUD operations', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const getOp = ops.find(
        (o) => o.docKind === DocOperationKind.GetResource,
      )!;
      expect(getOp.typespecOperation).toBeDefined();
      expect(getOp.typespecOperation!.kind).toBe('Operation');
    });

    it('populates typespecOperation for action operations', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @operationParameters model testActionParams {
            value: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            doSomething is GraphOps.Action<TActionParams=testActionParams, TReturnType=testItem>;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const actionOp = ops.find((o) => o.docKind === DocOperationKind.Action)!;
      expect(actionOp.typespecOperation).toBeDefined();
    });
  });

  describe('hidden operations', () => {
    it('skips operations with IsHidden attribute', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
            @agsAttribute("IsHidden", "true")
            patch is GraphOps.PatchNoResponse;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      expect(
        ops.find((o) => o.docKind === DocOperationKind.GetResource),
      ).toBeDefined();
      expect(
        ops.find((o) => o.docKind === DocOperationKind.Update),
      ).toBeUndefined();
    });
  });

  describe('parent entity resolution', () => {
    it('resolves parent segment from nested route', async () => {
      await runner.compile(`
        using MsGraph;
        @publicNamespace("microsoft.graph")
        namespace microsoft.graph {
          @entity model message {
            @readOnly @computed @key id: string;
          }
          @entity model user {
            @readOnly @computed @key id: string;
            @contains messages: message[];
          }
          @graphRoute("users")
          interface users extends Collection<user> {}
          @graphRoute("users/{userId}")
          interface usersById extends Resource<user> {}
          @graphRoute("users/{userId}/messages")
          interface userMessages extends Collection<message> {
            post is GraphOps.Post;
          }
        }
      `);

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('messages'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'message',
      );

      const postOp = ops.find(
        (o) => o.docKind === DocOperationKind.PostCreate,
      )!;
      expect(postOp.parentEntityName).toBe('users');
    });
  });
});
