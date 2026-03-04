// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { createGraphDocsTestRunner, graphSpec} from '../test-host.js';
import { collectGraphTypes } from '../../src/utils/type-collector.js';
import {
  formatTypeName,
  formatJsonValue,
} from '../../src/utils/type-formatter.js';

let runner: BasicTestRunner;

beforeEach(async () => {
  runner = await createGraphDocsTestRunner();
});

describe('formatTypeName', () => {
  describe('scalar types', () => {
    it('maps string to String', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: string; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('String');
    });

    it('maps boolean to Boolean', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: boolean; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('Boolean');
    });

    it('maps int64 to Int64', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: int64; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('Int64');
    });

    it('maps float64 to Double', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: float64; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('Double');
    });

    it('maps utcDateTime to DateTimeOffset', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: utcDateTime; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('DateTimeOffset');
    });

    it('maps bytes to Binary', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: bytes; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('Binary');
    });

    it('maps duration to Duration', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: duration; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('Duration');
    });

    it('maps plainDate to Date', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: plainDate; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('Date');
    });
  });

  describe('enum types', () => {
    it('renders enum as markdown link', async () => {
      await runner.compile(graphSpec(`
          enum testStatus { active, inactive, unknownFutureValue }
          @complex model testModel { status: testStatus; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('status')!;
      expect(formatTypeName(prop.type)).toBe('[testStatus](teststatus.md)');
    });
  });

  describe('model types', () => {
    it('renders named model as markdown link', async () => {
      await runner.compile(graphSpec(`
          @complex model innerModel { value: string; }
          @complex model testModel { inner: innerModel; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('inner')!;
      expect(formatTypeName(prop.type)).toBe('[innerModel](innermodel.md)');
    });

    it('renders array as collection type', async () => {
      await runner.compile(graphSpec(`
          @complex model innerModel { value: string; }
          @complex model testModel { items: innerModel[]; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('items')!;
      expect(formatTypeName(prop.type)).toBe(
        '[innerModel](innermodel.md) collection',
      );
    });

    it('renders string array as String collection', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { tags: string[]; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('tags')!;
      expect(formatTypeName(prop.type)).toBe('String collection');
    });
  });

  describe('union types', () => {
    it('unwraps nullable union to inner type', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: string | null; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('String');
    });

    it('joins multiple non-null variants with or', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: string | int32; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatTypeName(prop.type)).toBe('String or Int32');
    });
  });
});

describe('formatJsonValue', () => {
  describe('scalar types', () => {
    it('returns "String" for string', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: string; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatJsonValue(prop.type)).toBe('"String"');
    });

    it('returns true for boolean', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: boolean; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatJsonValue(prop.type)).toBe('true');
    });

    it('returns 0 for int32', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: int32; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatJsonValue(prop.type)).toBe('0');
    });

    it('returns 0.0 for float64', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: float64; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatJsonValue(prop.type)).toBe('0.0');
    });

    it('returns "String (timestamp)" for utcDateTime', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: utcDateTime; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatJsonValue(prop.type)).toBe('"String (timestamp)"');
    });

    it('returns "Duration" for duration', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: duration; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatJsonValue(prop.type)).toBe('"Duration"');
    });

    it('returns "Binary" for bytes', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: bytes; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatJsonValue(prop.type)).toBe('"Binary"');
    });

    it('returns "Date" for plainDate', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: plainDate; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatJsonValue(prop.type)).toBe('"Date"');
    });
  });

  describe('model types', () => {
    it('returns @odata.type for named model', async () => {
      await runner.compile(graphSpec(`
          @complex model innerModel { value: string; }
          @complex model testModel { inner: innerModel; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('inner')!;
      expect(formatJsonValue(prop.type)).toBe(
        '{"@odata.type": "microsoft.graph.innerModel"}',
      );
    });

    it('returns [] for array type', async () => {
      await runner.compile(graphSpec(`
          @complex model innerModel { value: string; }
          @complex model testModel { items: innerModel[]; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('items')!;
      expect(formatJsonValue(prop.type)).toBe('[]');
    });
  });

  describe('enum types', () => {
    it('returns "String" for enum', async () => {
      await runner.compile(graphSpec(`
          enum testStatus { active, inactive, unknownFutureValue }
          @complex model testModel { status: testStatus; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('status')!;
      expect(formatJsonValue(prop.type)).toBe('"String"');
    });
  });

  describe('union types', () => {
    it('unwraps nullable union for JSON value', async () => {
      await runner.compile(graphSpec(`
          @complex model testModel { value: string | null; }
      `));
      const types = collectGraphTypes(runner.program);
      const model = types.complexTypes.find((c) => c.name === 'testModel')!;
      const prop = model.model.properties.get('value')!;
      expect(formatJsonValue(prop.type)).toBe('"String"');
    });
  });
});
