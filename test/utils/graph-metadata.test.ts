// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { createGraphDocsTestRunner, graphSpec } from '../test-host.js';
import { collectGraphTypes } from '../../src/utils/type-collector.js';
import {
  classifyProperty,
  getDescription,
} from '../../src/utils/graph-metadata.js';

let runner: BasicTestRunner;

beforeEach(async () => {
  runner = await createGraphDocsTestRunner();
});

describe('classifyProperty', () => {
  it('identifies @computed property', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @computed displayName: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;
    const prop = entity.model.properties.get('displayName')!;
    const cls = classifyProperty(runner.program, prop);

    expect(cls.isComputed).toBe(true);
    expect(cls.isContainment).toBe(false);
  });

  it('identifies @readOnly property', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;
    const prop = entity.model.properties.get('id')!;
    const cls = classifyProperty(runner.program, prop);

    expect(cls.isReadOnly).toBe(true);
    expect(cls.isComputed).toBe(true);
  });

  it('identifies @immutable property', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @immutable @requiredForCreate tenantId: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;
    const prop = entity.model.properties.get('tenantId')!;
    const cls = classifyProperty(runner.program, prop);

    expect(cls.isImmutable).toBe(true);
    expect(cls.isRequiredForCreate).toBe(true);
  });

  it('identifies @requiredForCreate property', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @requiredForCreate displayName: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;
    const prop = entity.model.properties.get('displayName')!;
    const cls = classifyProperty(runner.program, prop);

    expect(cls.isRequiredForCreate).toBe(true);
    expect(cls.isComputed).toBe(false);
    expect(cls.isReadOnly).toBe(false);
    expect(cls.isImmutable).toBe(false);
  });

  it('identifies @contains (containment) property', async () => {
    await runner.compile(
      graphSpec(`
        @entity model child {
          @readOnly @computed @key id: string;
        }
        @entity model parent {
          @readOnly @computed @key id: string;
          @contains children: child[];
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'parent')!;
    const prop = entity.model.properties.get('children')!;
    const cls = classifyProperty(runner.program, prop);

    expect(cls.isContainment).toBe(true);
  });

  it('identifies nullable property (union with null)', async () => {
    await runner.compile(
      graphSpec(`
        @complex model testModel {
          value: string | null;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const model = types.complexTypes.find((c) => c.name === 'testModel')!;
    const prop = model.model.properties.get('value')!;
    const cls = classifyProperty(runner.program, prop);

    expect(cls.isNullable).toBe(true);
  });

  it('identifies non-nullable property', async () => {
    await runner.compile(
      graphSpec(`
        @complex model testModel {
          value: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const model = types.complexTypes.find((c) => c.name === 'testModel')!;
    const prop = model.model.properties.get('value')!;
    const cls = classifyProperty(runner.program, prop);

    expect(cls.isNullable).toBe(false);
  });

  it('identifies union without null as non-nullable', async () => {
    await runner.compile(
      graphSpec(`
        @complex model testModel {
          value: string | int32;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const model = types.complexTypes.find((c) => c.name === 'testModel')!;
    const prop = model.model.properties.get('value')!;
    const cls = classifyProperty(runner.program, prop);

    expect(cls.isNullable).toBe(false);
  });
});

describe('getDescription', () => {
  it('returns doc comment string', async () => {
    await runner.compile(
      graphSpec(`
        /** A test entity. */
        @entity model testEntity {
          @readOnly @computed @key id: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    expect(getDescription(runner.program, entity.model)).toBe('A test entity.');
  });

  it('returns undefined when no doc comment', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;
    const prop = entity.model.properties.get('id')!;

    expect(getDescription(runner.program, prop)).toBeUndefined();
  });

  it('returns property doc comment', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          /** The unique identifier. */
          @readOnly @computed @key id: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;
    const prop = entity.model.properties.get('id')!;

    expect(getDescription(runner.program, prop)).toBe('The unique identifier.');
  });
});
