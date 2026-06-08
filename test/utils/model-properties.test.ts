// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { createGraphDocsTestRunner, graphSpec } from '../test-host.js';
import { collectGraphTypes } from '../../src/utils/type-collector.js';
import { getModelProperties } from '../../src/utils/model-properties.js';

let runner: BasicTestRunner;

beforeEach(async () => {
  runner = await createGraphDocsTestRunner();
});

describe('getModelProperties', () => {
  it('returns own properties of a model', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          displayName: string;
          status: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;
    const props = getModelProperties(runner.program, entity.model);

    const names = props.map(([name]) => name);
    expect(names).toContain('id');
    expect(names).toContain('displayName');
    expect(names).toContain('status');
    expect(names).toHaveLength(3);
  });

  it('includes inherited base model properties', async () => {
    await runner.compile(
      graphSpec(`
        @entity model baseEntity {
          @readOnly @computed @key id: string;
          createdAt: string;
        }
        @entity model childEntity extends baseEntity {
          name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'childEntity')!;
    const props = getModelProperties(runner.program, entity.model);

    const names = props.map(([name]) => name);
    expect(names).toContain('name');
    expect(names).toContain('createdAt');
  });

  describe('containment filtering', () => {
    it('excludes @contains properties with "exclude" filter', async () => {
      await runner.compile(
        graphSpec(`
          @entity model child {
            @readOnly @computed @key id: string;
          }
          @entity model parent {
            @readOnly @computed @key id: string;
            displayName: string;
            @contains children: child[];
          }
      `),
      );

      const types = collectGraphTypes(runner.program);
      const entity = types.entities.find((e) => e.name === 'parent')!;
      const props = getModelProperties(runner.program, entity.model, 'exclude');

      const names = props.map(([name]) => name);
      expect(names).toContain('id');
      expect(names).toContain('displayName');
      expect(names).not.toContain('children');
    });

    it('returns only @contains properties with "only" filter', async () => {
      await runner.compile(
        graphSpec(`
          @entity model child {
            @readOnly @computed @key id: string;
          }
          @entity model parent {
            @readOnly @computed @key id: string;
            displayName: string;
            @contains children: child[];
          }
      `),
      );

      const types = collectGraphTypes(runner.program);
      const entity = types.entities.find((e) => e.name === 'parent')!;
      const props = getModelProperties(runner.program, entity.model, 'only');

      const names = props.map(([name]) => name);
      expect(names).toEqual(['children']);
    });

    it('returns all properties with "all" filter', async () => {
      await runner.compile(
        graphSpec(`
          @entity model child {
            @readOnly @computed @key id: string;
          }
          @entity model parent {
            @readOnly @computed @key id: string;
            displayName: string;
            @contains children: child[];
          }
      `),
      );

      const types = collectGraphTypes(runner.program);
      const entity = types.entities.find((e) => e.name === 'parent')!;
      const props = getModelProperties(runner.program, entity.model, 'all');

      const names = props.map(([name]) => name);
      expect(names).toContain('id');
      expect(names).toContain('displayName');
      expect(names).toContain('children');
      expect(names).toHaveLength(3);
    });

    it('defaults to "all" when no filter specified', async () => {
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
      const propsDefault = getModelProperties(runner.program, entity.model);
      const propsAll = getModelProperties(runner.program, entity.model, 'all');

      expect(propsDefault.map(([n]) => n)).toEqual(propsAll.map(([n]) => n));
    });
  });

  it('returns empty array for model with no properties', async () => {
    await runner.compile(
      graphSpec(`
        @complex model emptyModel {
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const complex = types.complexTypes.find((c) => c.name === 'emptyModel')!;
    const props = getModelProperties(runner.program, complex.model);

    expect(props).toHaveLength(0);
  });

  it('filters inherited @contains properties correctly', async () => {
    await runner.compile(
      graphSpec(`
        @entity model grandchild {
          @readOnly @computed @key id: string;
        }
        @entity model baseWithNav {
          @readOnly @computed @key id: string;
          @contains items: grandchild[];
        }
        @entity model derived extends baseWithNav {
          name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'derived')!;

    const excluded = getModelProperties(
      runner.program,
      entity.model,
      'exclude',
    );
    const excludedNames = excluded.map(([name]) => name);
    expect(excludedNames).toContain('id');
    expect(excludedNames).toContain('name');
    expect(excludedNames).not.toContain('items');

    const navOnly = getModelProperties(runner.program, entity.model, 'only');
    const navNames = navOnly.map(([name]) => name);
    expect(navNames).toEqual(['items']);
  });

  it('walks multi-level inheritance chains', async () => {
    await runner.compile(
      graphSpec(`
        @entity model grandparent {
          @readOnly @computed @key id: string;
          createdAt: string;
        }
        @entity model parent extends grandparent {
          updatedAt: string;
        }
        @entity model child extends parent {
          name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'child')!;
    const props = getModelProperties(runner.program, entity.model);

    const names = props.map(([name]) => name);
    expect(names).toContain('name');
    expect(names).toContain('updatedAt');
    expect(names).toContain('createdAt');
    expect(names).toContain('id');
  });

  it('derived property overrides base property with same name', async () => {
    await runner.compile(
      graphSpec(`
        @entity model base {
          @readOnly @computed @key id: string;
          /** Base description */
          status: string;
        }
        @entity model derived extends base {
          /** Derived description */
          status: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'derived')!;
    const props = getModelProperties(runner.program, entity.model);

    // "status" should appear only once (derived wins)
    const statusProps = props.filter(([name]) => name === 'status');
    expect(statusProps).toHaveLength(1);
  });
});
