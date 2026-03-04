// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore msgraph testitem onEmit

import { describe, it, expect, beforeAll } from 'vitest';
import { resolvePath } from '@typespec/compiler';
import { createTester } from '@typespec/compiler/testing';
import type {
  EmitterTester,
  TestEmitterCompileResult,
} from '@typespec/compiler/testing';

const EMITTER_NAME = '@microsoft/typespec-graph-docs-emitter';
const EMITTER_BASE = resolvePath(import.meta.dirname, '../..');
const TESTER_LIBRARIES = [
  '@microsoft/typespec-msgraph',
  '@typespec/http',
  EMITTER_NAME,
];

let emitterTester: EmitterTester;

beforeAll(() => {
  emitterTester = createTester(EMITTER_BASE, {
    libraries: TESTER_LIBRARIES,
  })
    .importLibraries()
    .emit(EMITTER_NAME);
});

const BASIC_SPEC = `
using MsGraph;

@publicNamespace("microsoft.graph")
namespace microsoft.graph {
  /** A test entity. */
  @entity model testItem {
    /** The unique identifier. */
    @readOnly @computed @key id: string;
    /** The display name. */
    @computed displayName: string;
  }

  @graphRoute("items")
  interface testItems extends Collection<testItem> {
    post is GraphOps.Post;
  }

  @graphRoute("items/{id}")
  interface testItemsById extends Resource<testItem> {
    get is GraphOps.GetResource;
    patch is GraphOps.PatchNoResponse;
    delete is GraphOps.Delete;
  }
}
`;

const SPEC_WITH_ENUMS = `
using MsGraph;

@publicNamespace("microsoft.graph")
namespace microsoft.graph {
  enum testStatus { active, inactive, unknownFutureValue }

  /** An entity. */
  @entity model testItem {
    /** The unique identifier. */
    @readOnly @computed @key id: string;
    /** The status. */
    @computed status: testStatus;
  }
}
`;

const SPEC_WITH_COMPLEX = `
using MsGraph;

@publicNamespace("microsoft.graph")
namespace microsoft.graph {
  /** A complex type. */
  @complex model testAddress {
    /** The street. */
    street: string;
    /** The city. */
    city: string;
  }

  /** An entity. */
  @entity model testItem {
    /** The unique identifier. */
    @readOnly @computed @key id: string;
    /** The mailing address. */
    @computed address: testAddress;
  }
}
`;

const SPEC_NO_ENUMS = `
using MsGraph;

@publicNamespace("microsoft.graph")
namespace microsoft.graph {
  /** An entity. */
  @entity model testItem {
    /** The unique identifier. */
    @readOnly @computed @key id: string;
    /** The name. */
    @computed displayName: string;
  }
}
`;

describe('$onEmit end-to-end', () => {
  describe('file output structure', () => {
    let result: TestEmitterCompileResult;

    beforeAll(async () => {
      result = await emitterTester.compile(BASIC_SPEC);
    });

    it('creates resource type file', () => {
      expect(result.outputs['v1.0/resources/testitem.md']).toBeDefined();
    });

    it('creates GET method file', () => {
      expect(result.outputs['v1.0/api/testitem-get.md']).toBeDefined();
    });

    it('creates PATCH method file', () => {
      expect(result.outputs['v1.0/api/testitem-update.md']).toBeDefined();
    });

    it('creates DELETE method file', () => {
      expect(result.outputs['v1.0/api/testitem-delete.md']).toBeDefined();
    });

    it('creates POST method file', () => {
      // POST on a collection -> {parent}-post-{plural}
      // Since this is a root collection without a parent entity, check filename
      const apiFiles = Object.keys(result.outputs).filter((k) =>
        k.startsWith('v1.0/api/'),
      );
      const postFile = apiFiles.find((f) => f.includes('post'));
      expect(postFile).toBeDefined();
    });
  });

  describe('resource type page content', () => {
    let result: TestEmitterCompileResult;

    beforeAll(async () => {
      result = await emitterTester.compile(BASIC_SPEC);
    });

    it('contains YAML front matter', () => {
      const content = result.outputs['v1.0/resources/testitem.md'];
      expect(content).toContain('---');
      expect(content).toContain('title: "testItem resource type"');
      expect(content).toContain('doc_type: resourcePageType');
    });

    it('contains heading and namespace', () => {
      const content = result.outputs['v1.0/resources/testitem.md'];
      expect(content).toContain('# testItem resource type');
      expect(content).toContain('Namespace: microsoft.graph');
    });

    it('contains description', () => {
      const content = result.outputs['v1.0/resources/testitem.md'];
      expect(content).toContain('A test entity.');
    });

    it('contains properties table', () => {
      const content = result.outputs['v1.0/resources/testitem.md'];
      expect(content).toContain('## Properties');
      expect(content).toContain('| id | String |');
      expect(content).toContain('| displayName | String |');
    });

    it('contains methods table', () => {
      const content = result.outputs['v1.0/resources/testitem.md'];
      expect(content).toContain('## Methods');
    });

    it('contains JSON representation', () => {
      const content = result.outputs['v1.0/resources/testitem.md'];
      expect(content).toContain('## JSON representation');
      expect(content).toContain('"@odata.type": "#microsoft.graph.testItem"');
    });
  });

  describe('API method page content', () => {
    let result: TestEmitterCompileResult;

    beforeAll(async () => {
      result = await emitterTester.compile(BASIC_SPEC);
    });

    it('GET page has correct structure', () => {
      const content = result.outputs['v1.0/api/testitem-get.md'];
      expect(content).toContain('doc_type: apiPageType');
      expect(content).toContain('## HTTP request');
      expect(content).toContain('GET');
      expect(content).toContain('## Response');
    });

    it('PATCH page has request body section', () => {
      const content = result.outputs['v1.0/api/testitem-update.md'];
      expect(content).toContain('PATCH');
      expect(content).toContain('## Request body');
    });

    it('DELETE page has correct HTTP method', () => {
      const content = result.outputs['v1.0/api/testitem-delete.md'];
      expect(content).toContain('DELETE');
    });
  });

  describe('enum page conditional', () => {
    it('creates enums.md when enums exist', async () => {
      const result = await emitterTester.compile(SPEC_WITH_ENUMS);
      expect(result.outputs['v1.0/resources/enums.md']).toBeDefined();
      const content = result.outputs['v1.0/resources/enums.md'];
      expect(content).toContain('### testStatus values');
      expect(content).toContain('| active |');
    });

    it('does not create enums.md when no enums exist', async () => {
      const result = await emitterTester.compile(SPEC_NO_ENUMS);
      expect(result.outputs['v1.0/resources/enums.md']).toBeUndefined();
    });
  });

  describe('complex type pages', () => {
    it('creates complex type file', async () => {
      const result = await emitterTester.compile(SPEC_WITH_COMPLEX);
      expect(result.outputs['v1.0/resources/testaddress.md']).toBeDefined();
      const content = result.outputs['v1.0/resources/testaddress.md'];
      expect(content).toContain('# testAddress resource type');
      expect(content).toContain('| street | String |');
      expect(content).toContain('| city | String |');
    });
  });

  describe('emitter options', () => {
    it('uses api-version option in output path', async () => {
      const betaTester = createTester(EMITTER_BASE, {
        libraries: TESTER_LIBRARIES,
      })
        .importLibraries()
        .emit(EMITTER_NAME, { 'api-version': 'beta' });

      const result = await betaTester.compile(SPEC_NO_ENUMS);
      expect(result.outputs['beta/resources/testitem.md']).toBeDefined();
      const content = result.outputs['beta/resources/testitem.md'];
      expect(content).toContain('[!INCLUDE [beta-disclaimer');
    });

    it('uses ms-date option in front matter', async () => {
      const dateTester = createTester(EMITTER_BASE, {
        libraries: TESTER_LIBRARIES,
      })
        .importLibraries()
        .emit(EMITTER_NAME, { 'ms-date': '03/15/2025' });

      const result = await dateTester.compile(SPEC_NO_ENUMS);
      const content = result.outputs['v1.0/resources/testitem.md'];
      expect(content).toContain('ms.date: 03/15/2025');
    });

    it('uses author option in front matter', async () => {
      const authorTester = createTester(EMITTER_BASE, {
        libraries: TESTER_LIBRARIES,
      })
        .importLibraries()
        .emit(EMITTER_NAME, { author: 'testuser' });

      const result = await authorTester.compile(SPEC_NO_ENUMS);
      const content = result.outputs['v1.0/resources/testitem.md'];
      expect(content).toContain('author: "testuser"');
    });

    it('defaults api-version to v1.0', async () => {
      const result = await emitterTester.compile(SPEC_NO_ENUMS);
      const keys = Object.keys(result.outputs);
      expect(keys.every((k) => k.startsWith('v1.0/'))).toBe(true);
    });
  });

  describe('operation aggregation', () => {
    it('aggregates operations from multiple routes for same entity', async () => {
      const result = await emitterTester.compile(BASIC_SPEC);
      const content = result.outputs['v1.0/resources/testitem.md'];
      // Methods table should contain links to get, update, delete, post pages
      expect(content).toContain('## Methods');
      // Count the method links in the Methods section
      const methodsSection = content.split('## Methods')[1]?.split('## ')[0];
      expect(methodsSection).toBeDefined();
      // Should have multiple method rows
      const methodRows = methodsSection!
        .split('\n')
        .filter((l: string) => l.startsWith('| ['));
      expect(methodRows.length).toBeGreaterThanOrEqual(3);
    });
  });
});
