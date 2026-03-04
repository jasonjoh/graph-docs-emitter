// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { createGraphDocsTestRunner, graphSpec} from '../test-host.js';
import {
  $exampleRequest,
  getExampleRequest,
} from '../../src/decorators/example-request.js';
import {
  $exampleResponse,
  getExampleResponse,
} from '../../src/decorators/example-response.js';

let runner: BasicTestRunner;

beforeEach(async () => {
  runner = await createGraphDocsTestRunner();
});

describe('$exampleRequest / getExampleRequest', () => {
  it('stores and retrieves an object value', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }

        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
        }
    `));

    const iface = runner.program
      .getGlobalNamespaceType()
      .namespaces.get('microsoft')
      ?.namespaces.get('graph')
      ?.interfaces.get('testItemsById');
    const op = iface?.operations.get('get');
    expect(op).toBeDefined();

    const exampleObj = { displayName: 'Test Widget', count: 42 };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    $exampleRequest({ program: runner.program } as any, op!, exampleObj);

    const result = getExampleRequest(runner.program, op!);
    expect(result).toEqual({ displayName: 'Test Widget', count: 42 });
  });

  it('returns undefined when no decorator applied', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }

        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
        }
    `));

    const iface = runner.program
      .getGlobalNamespaceType()
      .namespaces.get('microsoft')
      ?.namespaces.get('graph')
      ?.interfaces.get('testItemsById');
    const op = iface?.operations.get('get');
    expect(op).toBeDefined();

    const result = getExampleRequest(runner.program, op!);
    expect(result).toBeUndefined();
  });

  it('stores a string value', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }

        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
        }
    `));

    const iface = runner.program
      .getGlobalNamespaceType()
      .namespaces.get('microsoft')
      ?.namespaces.get('graph')
      ?.interfaces.get('testItemsById');
    const op = iface?.operations.get('get');

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    $exampleRequest({ program: runner.program } as any, op!, 'raw string body');

    expect(getExampleRequest(runner.program, op!)).toBe('raw string body');
  });

  it('stores an array value', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }

        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
        }
    `));

    const iface = runner.program
      .getGlobalNamespaceType()
      .namespaces.get('microsoft')
      ?.namespaces.get('graph')
      ?.interfaces.get('testItemsById');
    const op = iface?.operations.get('get');

    const arr = [{ id: '1' }, { id: '2' }];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    $exampleRequest({ program: runner.program } as any, op!, arr);

    expect(getExampleRequest(runner.program, op!)).toEqual([
      { id: '1' },
      { id: '2' },
    ]);
  });
});

describe('$exampleResponse / getExampleResponse', () => {
  it('stores and retrieves an object value', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }

        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
        }
    `));

    const iface = runner.program
      .getGlobalNamespaceType()
      .namespaces.get('microsoft')
      ?.namespaces.get('graph')
      ?.interfaces.get('testItemsById');
    const op = iface?.operations.get('get');
    expect(op).toBeDefined();

    const exampleObj = {
      id: 'abc-123',
      displayName: 'Test Widget',
      state: 'active',
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    $exampleResponse({ program: runner.program } as any, op!, exampleObj);

    const result = getExampleResponse(runner.program, op!);
    expect(result).toEqual({
      id: 'abc-123',
      displayName: 'Test Widget',
      state: 'active',
    });
  });

  it('returns undefined when no decorator applied', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }

        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
        }
    `));

    const iface = runner.program
      .getGlobalNamespaceType()
      .namespaces.get('microsoft')
      ?.namespaces.get('graph')
      ?.interfaces.get('testItemsById');
    const op = iface?.operations.get('get');

    expect(getExampleResponse(runner.program, op!)).toBeUndefined();
  });

  it('stores different values for different operations', async () => {
    await runner.compile(graphSpec(`
        @entity model testItem {
          @readOnly @computed @key id: string;
        }

        @graphRoute("items/{id}")
        interface testItemsById extends Resource<testItem> {
          get is GraphOps.GetResource;
          patch is GraphOps.PatchNoResponse;
        }
    `));

    const iface = runner.program
      .getGlobalNamespaceType()
      .namespaces.get('microsoft')
      ?.namespaces.get('graph')
      ?.interfaces.get('testItemsById');
    const getOp = iface?.operations.get('get');
    const patchOp = iface?.operations.get('patch');

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    $exampleResponse({ program: runner.program } as any, getOp!, {
      id: 'get-response',
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    $exampleResponse({ program: runner.program } as any, patchOp!, {
      id: 'patch-response',
    });

    expect(getExampleResponse(runner.program, getOp!)).toEqual({
      id: 'get-response',
    });
    expect(getExampleResponse(runner.program, patchOp!)).toEqual({
      id: 'patch-response',
    });
  });
});
