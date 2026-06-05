// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { createGraphDocsTestRunner, graphSpec } from '../test-host.js';
import { collectGraphTypes } from '../../src/utils/type-collector.js';
import {
  resolveOperationsFromRoute,
  resolveGlobalOperations,
  buildEntityRouteIndex,
  DocOperationKind,
  getStandardCrudDescription,
} from '../../src/utils/operation-resolver.js';
import { GraphRouteInfo } from '../../src/utils/type-collector.js';

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items")
          interface testItems extends Collection<testItem> {
            @select @top @skip getAll is GraphOps.GetPagedCollection;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
            @computed displayName: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            compute is GraphOps.Function<TReturnType=testItem>;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            delete is GraphOps.Delete;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items")
          interface testItems extends Collection<testItem> {
            @select getAll is GraphOps.GetPagedCollection;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
            patch is GraphOps.PatchNoResponse;
            delete is GraphOps.Delete;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
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
      `),
      );

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
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
      `),
      );

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
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
            @agsAttribute("IsHidden", "true")
            patch is GraphOps.PatchNoResponse;
          }
      `),
      );

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
      await runner.compile(
        graphSpec(`
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
      `),
      );

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

  describe('Put operations', () => {
    it('resolves PutWithResponse as Update docKind with PUT method', async () => {
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
            name: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            put is GraphOps.PutWithResponse;
          }
      `),
      );

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const putOp = ops.find((o) => o.docKind === DocOperationKind.Update)!;
      expect(putOp).toBeDefined();
      expect(putOp.httpMethod).toBe('PUT');
      expect(putOp.returnTypeName).toBe('testItem');
    });
  });

  describe('deduplication', () => {
    it('does not produce duplicate operations from source interfaces', async () => {
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is GraphOps.GetResource;
            patch is GraphOps.PatchNoResponse;
            delete is GraphOps.Delete;
          }
      `),
      );

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('items'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      // Each operation kind should appear exactly once
      const getOps = ops.filter(
        (o) => o.docKind === DocOperationKind.GetResource,
      );
      const updateOps = ops.filter(
        (o) => o.docKind === DocOperationKind.Update,
      );
      const deleteOps = ops.filter(
        (o) => o.docKind === DocOperationKind.Delete,
      );
      expect(getOps).toHaveLength(1);
      expect(updateOps).toHaveLength(1);
      expect(deleteOps).toHaveLength(1);
    });
  });

  describe('get alias disambiguation', () => {
    it('resolves get alias as GetResource on single-resource route', async () => {
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }

          interface LocalOps {
            get is GraphOps.GetResource;
          }

          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            get is LocalOps.get;
          }
      `),
      );

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('{id}'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const getOp = ops.find((o) => o.httpMethod === 'GET');
      expect(getOp).toBeDefined();
      expect(getOp!.docKind).toBe(DocOperationKind.GetResource);
    });
  });

  describe('requestBodyModel filtering', () => {
    it('excludes @path parameters from requestBodyModel', async () => {
      await runner.compile(
        graphSpec(`
          @entity model testItem {
            @readOnly @computed @key id: string;
          }
          @operationParameters model doThingParams {
            /** The reason. */
            reason: string;
          }
          @graphRoute("items/{id}")
          interface testItemsById extends Resource<testItem> {
            doThing is GraphOps.Action<TActionParams=doThingParams>;
          }
      `),
      );

      const types = collectGraphTypes(runner.program);
      const route = types.routes.find((r) => r.path.includes('{id}'))!;
      const ops = resolveOperationsFromRoute(
        runner.program,
        route,
        types.entities,
        'testItem',
      );

      const actionOp = ops.find((o) => o.docKind === DocOperationKind.Action)!;
      expect(actionOp.requestBodyModel).toBeDefined();

      // requestBodyModel should only contain body params, not route params
      const paramNames = [...actionOp.requestBodyModel!.properties.keys()];
      expect(paramNames).toContain('reason');
      expect(paramNames).not.toContain('graphRouteParams');
    });
  });
});

describe('resolveGlobalOperations', () => {
  it('resolves @globalOperation actions', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }
        @operationParameters model resetParams {
          reason: string;
        }
        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
        }
        interface testItemActions extends Resource<testItem> {
          @globalOperation resetState is GraphOps.Action<TActionParams=resetParams, TReturnType=testItem>;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    expect(types.globalOperationInterfaces.length).toBeGreaterThan(0);

    const globalIface = types.globalOperationInterfaces.find(
      (g) => g.entityName === 'testItem',
    )!;
    expect(globalIface).toBeDefined();

    const route = types.routes.find((r) => r.path.includes('items'))!;
    const ops = resolveGlobalOperations(
      runner.program,
      globalIface,
      route.path,
    );

    expect(ops.length).toBeGreaterThan(0);
    const actionOp = ops.find((o) => o.docKind === DocOperationKind.Action);
    expect(actionOp).toBeDefined();
    expect(actionOp!.actionOrFunctionName).toBe('resetState');
    expect(actionOp!.httpMethod).toBe('POST');
  });

  it('skips operations not decorated with @globalOperation', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }
        @operationParameters model actionParams {
          value: string;
        }
        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
        }
        interface testItemActions extends Resource<testItem> {
          @globalOperation doThis is GraphOps.Action<TActionParams=actionParams, TReturnType=testItem>;
          doThat is GraphOps.Action<TActionParams=actionParams, TReturnType=testItem>;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const globalIface = types.globalOperationInterfaces.find(
      (g) => g.entityName === 'testItem',
    )!;

    const route = types.routes.find((r) => r.path.includes('items'))!;
    const ops = resolveGlobalOperations(
      runner.program,
      globalIface,
      route.path,
    );

    // Only the @globalOperation op should be resolved
    const opNames = ops.map((o) => o.actionOrFunctionName);
    expect(opNames).toContain('doThis');
    expect(opNames).not.toContain('doThat');
  });
});

describe('buildEntityRouteIndex', () => {
  it('indexes collection and resource routes by entity name', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }
        @graphRoute("items")
        interface testItems extends Collection<testItem> {}
        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {}
    `),
    );

    const types = collectGraphTypes(runner.program);
    const getEntityName = (route: GraphRouteInfo) => {
      for (const src of route.iface.sourceInterfaces) {
        if (src.templateMapper?.args) {
          for (const arg of src.templateMapper.args) {
            if (arg.entityKind === 'Type' && arg.kind === 'Model' && arg.name) {
              return arg.name;
            }
          }
        }
      }
      return undefined;
    };

    const index = buildEntityRouteIndex(types.routes, getEntityName);

    expect(index.has('testItem')).toBe(true);
    const entry = index.get('testItem')!;
    expect(entry.collectionRoutes).toContain('items');
    expect(entry.resourceRoutes.some((r) => r.includes('{id}'))).toBe(true);
  });

  it('produces empty index when getEntityName returns undefined for all routes', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }
        @graphRoute("items")
        interface testItems extends Collection<testItem> {}
    `),
    );

    const types = collectGraphTypes(runner.program);
    // Simulate entity name resolution failing for all routes
    const index = buildEntityRouteIndex(types.routes, () => undefined);

    expect(index.size).toBe(0);
  });

  it('handles multiple routes for the same entity', async () => {
    await runner.compile(
      graphSpec(`
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
        interface userMessages extends Collection<message> {}
        @graphRoute("users/{userId}/messages/{messageId}")
        interface userMessagesById extends Resource<message> {}
    `),
    );

    const types = collectGraphTypes(runner.program);
    const getEntityName = (route: GraphRouteInfo) => {
      for (const src of route.iface.sourceInterfaces) {
        if (src.templateMapper?.args) {
          for (const arg of src.templateMapper.args) {
            if (arg.entityKind === 'Type' && arg.kind === 'Model' && arg.name) {
              return arg.name;
            }
          }
        }
      }
      return undefined;
    };

    const index = buildEntityRouteIndex(types.routes, getEntityName);

    const entry = index.get('message')!;
    expect(entry).toBeDefined();
    expect(entry.collectionRoutes.length).toBeGreaterThanOrEqual(1);
    expect(entry.resourceRoutes.length).toBeGreaterThanOrEqual(1);
  });
});
