// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph

import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { createGraphDocsTestRunner } from '../test-host.js';
import {
  collectGraphTypes,
  normalizeDescription,
} from '../../src/utils/type-collector.js';

let runner: BasicTestRunner;

beforeEach(async () => {
  runner = await createGraphDocsTestRunner();
});

describe('normalizeDescription', () => {
  it('inserts space after period before uppercase', () => {
    expect(normalizeDescription('First sentence.Second sentence.')).toBe(
      'First sentence. Second sentence.',
    );
  });

  it('inserts space after exclamation before uppercase', () => {
    expect(normalizeDescription('Done!Next step.')).toBe('Done! Next step.');
  });

  it('inserts space after question mark before uppercase', () => {
    expect(normalizeDescription('Ready?Go now.')).toBe('Ready? Go now.');
  });

  it('does not modify already-spaced text', () => {
    expect(normalizeDescription('First. Second.')).toBe('First. Second.');
  });

  it('does not modify punctuation followed by lowercase', () => {
    expect(normalizeDescription('e.g. this')).toBe('e.g. this');
  });

  it('returns undefined for undefined input', () => {
    expect(normalizeDescription(undefined)).toBeUndefined();
  });

  it('returns empty string for empty input', () => {
    expect(normalizeDescription('')).toBe('');
  });

  it('handles multiple occurrences', () => {
    expect(normalizeDescription('A.B.C.Done.')).toBe('A. B. C. Done.');
  });
});

describe('collectGraphTypes - isHidden', () => {
  it('skips models with @agsAttribute("IsHidden", "true")', async () => {
    await runner.compile(`
      using MsGraph;
      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        @entity model visibleEntity {
          @readOnly @computed @key id: string;
        }
        @agsAttribute("IsHidden", "true")
        @entity model hiddenEntity {
          @readOnly @computed @key id: string;
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    expect(
      types.entities.find((e) => e.name === 'visibleEntity'),
    ).toBeDefined();
    expect(
      types.entities.find((e) => e.name === 'hiddenEntity'),
    ).toBeUndefined();
  });

  it('skips enums with @agsAttribute("IsHidden", "true")', async () => {
    await runner.compile(`
      using MsGraph;
      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        enum visibleEnum { a, b, unknownFutureValue }
        @agsAttribute("IsHidden", "true")
        enum hiddenEnum { x, y, unknownFutureValue }
      }
    `);

    const types = collectGraphTypes(runner.program);
    expect(types.enums.find((e) => e.name === 'visibleEnum')).toBeDefined();
    expect(types.enums.find((e) => e.name === 'hiddenEnum')).toBeUndefined();
  });

  it('skips interfaces with @agsAttribute("IsHidden", "true")', async () => {
    await runner.compile(`
      using MsGraph;
      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        @entity model testItem {
          @readOnly @computed @key id: string;
        }
        @graphRoute("items")
        interface visibleRoute extends Collection<testItem> {}
        @agsAttribute("IsHidden", "true")
        @graphRoute("secret")
        interface hiddenRoute extends Collection<testItem> {}
      }
    `);

    const types = collectGraphTypes(runner.program);
    expect(types.routes.find((r) => r.path === 'items')).toBeDefined();
    expect(types.routes.find((r) => r.path === 'secret')).toBeUndefined();
  });
});

describe('collectGraphTypes - complex types', () => {
  it('collects @complex models', async () => {
    await runner.compile(`
      using MsGraph;
      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        /** A complex type. */
        @complex model testComplex {
          value: string;
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    const complex = types.complexTypes.find((c) => c.name === 'testComplex');
    expect(complex).toBeDefined();
    expect(complex!.description).toBe('A complex type.');
  });
});

describe('collectGraphTypes - description normalization', () => {
  it('normalizes concatenated descriptions on models', async () => {
    await runner.compile(`
      using MsGraph;
      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        /** First sentence.Second sentence. */
        @entity model testEntity {
          @readOnly @computed @key id: string;
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;
    expect(entity.description).toBe('First sentence. Second sentence.');
  });

  it('normalizes concatenated descriptions on enums', async () => {
    await runner.compile(`
      using MsGraph;
      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        /** An enum.Values listed below. */
        enum testEnum { a, unknownFutureValue }
      }
    `);

    const types = collectGraphTypes(runner.program);
    const enumInfo = types.enums.find((e) => e.name === 'testEnum')!;
    expect(enumInfo.description).toBe('An enum. Values listed below.');
  });
});

describe('collectGraphTypes - @operationParameters skip', () => {
  it('skips models with @operationParameters', async () => {
    await runner.compile(`
      using MsGraph;
      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        @operationParameters model testParams {
          value: string;
        }
        @entity model testEntity {
          @readOnly @computed @key id: string;
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    expect(types.entities.find((e) => e.name === 'testParams')).toBeUndefined();
    expect(
      types.complexTypes.find((c) => c.name === 'testParams'),
    ).toBeUndefined();
    expect(types.entities.find((e) => e.name === 'testEntity')).toBeDefined();
  });
});

describe('collectGraphTypes - namespace filtering', () => {
  it('skips types outside @publicNamespace', async () => {
    const [, diagnostics] = await runner.compileAndDiagnose(`
      using MsGraph;
      namespace myPrivate {
        @entity model hiddenEntity {
          @readOnly @computed @key id: string;
        }
      }
    `);

    // There will be diagnostics about missing @publicNamespace
    expect(diagnostics.length).toBeGreaterThan(0);
    const types = collectGraphTypes(runner.program);
    expect(
      types.entities.find((e) => e.name === 'hiddenEntity'),
    ).toBeUndefined();
  });
});
