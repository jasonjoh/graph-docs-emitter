// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { createGraphDocsTestRunner, graphSpec} from '../test-host.js';
import { collectGraphTypes } from '../../src/utils/type-collector.js';
import {
  resolveOperationsFromRoute,
  DocOperationKind,
} from '../../src/utils/operation-resolver.js';

let runner: BasicTestRunner;

beforeEach(async () => {
  runner = await createGraphDocsTestRunner();
});

describe('type-collector', () => {
  it('collects entities annotated with @entity', async () => {
    await runner.compile(graphSpec(`
        /** A test entity. */
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @computed displayName: string;
        }
    `));

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity');
    expect(entity).toBeDefined();
    expect(entity!.description).toBe('A test entity.');
  });

  it('collects complex types annotated with @complex', async () => {
    await runner.compile(graphSpec(`
        /** A complex type. */
        @complex model testComplex {
          value: string;
        }
    `));

    const types = collectGraphTypes(runner.program);
    const complex = types.complexTypes.find((c) => c.name === 'testComplex');
    expect(complex).toBeDefined();
    expect(complex!.description).toBe('A complex type.');
  });

  it('collects enums in public namespaces', async () => {
    await runner.compile(graphSpec(`
        /** The status. */
        enum testStatus {
          active,
          inactive,
          unknownFutureValue
        }
    `));

    const types = collectGraphTypes(runner.program);
    const enumInfo = types.enums.find((e) => e.name === 'testStatus');
    expect(enumInfo).toBeDefined();
    expect(enumInfo!.description).toBe('The status.');
    expect(enumInfo!.enumType.members.size).toBe(3);
  });

  it('collects routes from interfaces with @graphRoute', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
          @computed displayName: string;
        }

        @graphRoute("items")
        interface testItems extends Collection<testItem> {
        }
    `));

    const types = collectGraphTypes(runner.program);
    const route = types.routes.find((r) => r.path === 'items');
    expect(route).toBeDefined();
  });

  it('skips @operationParameters models', async () => {
    await runner.compile(graphSpec(`
        @operationParameters model testParams {
          value: string;
        }
    `));

    const types = collectGraphTypes(runner.program);
    expect(types.entities.find((e) => e.name === 'testParams')).toBeUndefined();
    expect(
      types.complexTypes.find((c) => c.name === 'testParams'),
    ).toBeUndefined();
  });

  it('skips types outside @publicNamespace', async () => {
    const [_, diagnostics] = await runner.compileAndDiagnose(`
      using MsGraph;

      namespace myPrivate {
        @entity model hiddenEntity {
          @readOnly @computed @key id: string;
        }
      }
    `);

    // Expect msgraph diagnostics about missing @publicNamespace, but still check the result
    expect(diagnostics.length).toBeGreaterThan(0);
    const types = collectGraphTypes(runner.program);
    expect(
      types.entities.find((e) => e.name === 'hiddenEntity'),
    ).toBeUndefined();
  });
});

describe('operation-resolver', () => {
  it('resolves operations from Collection interface with post', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
          @computed displayName: string;
        }

        @graphRoute("items")
        interface testItems extends Collection<testItem> {
          post is GraphOps.Post;
        }
    `));

    const types = collectGraphTypes(runner.program);
    const route = types.routes.find((r) => r.path === 'items');
    expect(route).toBeDefined();

    const ops = resolveOperationsFromRoute(
      runner.program,
      route!,
      types.entities,
    );
    const postOp = ops.find((o) => o.docKind === DocOperationKind.PostCreate);
    expect(postOp).toBeDefined();
    expect(postOp!.httpMethod).toBe('POST');
  });

  it('resolves Resource interface operations', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
          @computed displayName: string;
        }

        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
          patch is GraphOps.PatchNoResponse;
          delete is GraphOps.Delete;
        }
    `));

    const types = collectGraphTypes(runner.program);
    const route = types.routes.find((r) => r.path.includes('items'));
    expect(route).toBeDefined();

    const ops = resolveOperationsFromRoute(
      runner.program,
      route!,
      types.entities,
    );

    const getOp = ops.find((o) => o.docKind === DocOperationKind.GetResource);
    expect(getOp).toBeDefined();
    expect(getOp!.httpMethod).toBe('GET');

    const patchOp = ops.find((o) => o.docKind === DocOperationKind.Update);
    expect(patchOp).toBeDefined();
    expect(patchOp!.httpMethod).toBe('PATCH');

    const deleteOp = ops.find((o) => o.docKind === DocOperationKind.Delete);
    expect(deleteOp).toBeDefined();
    expect(deleteOp!.httpMethod).toBe('DELETE');
  });

  it('resolves action operations', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
          @computed displayName: string;
        }

        @operationParameters model testActionParams {
          value: string;
        }

        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          doSomething is GraphOps.Action<TActionParams=testActionParams, TReturnType=testItem>;
        }
    `));

    const types = collectGraphTypes(runner.program);
    const route = types.routes.find((r) => r.path.includes('items'));
    expect(route).toBeDefined();

    const ops = resolveOperationsFromRoute(
      runner.program,
      route!,
      types.entities,
    );
    const actionOp = ops.find((o) => o.docKind === DocOperationKind.Action);
    expect(actionOp).toBeDefined();
    expect(actionOp!.actionOrFunctionName).toBe('doSomething');
    expect(actionOp!.httpMethod).toBe('POST');
  });
});
