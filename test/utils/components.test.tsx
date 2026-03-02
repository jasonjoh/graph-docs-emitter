// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore jasonjoh testitem

/** @jsxImportSource @alloy-js/core */
import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { renderTree, printTree } from '@alloy-js/core';
import { createGraphDocsTestRunner } from '../test-host.js';
import { collectGraphTypes } from '../../src/utils/type-collector.js';
import { PropertiesTable } from '../../src/components/PropertiesTable.jsx';
import { RelationshipsTable } from '../../src/components/RelationshipsTable.jsx';
import { EnumMembersTable } from '../../src/components/EnumMembersTable.jsx';
import { JsonRepresentation } from '../../src/components/JsonRepresentation.jsx';
import { YamlFrontMatter } from '../../src/components/YamlFrontMatter.jsx';
import { MethodsTable } from '../../src/components/MethodsTable.jsx';
import { DocOperationKind } from '../../src/utils/operation-resolver.js';

function renderToString(jsx: unknown): string {
  const tree = renderTree(jsx as Parameters<typeof renderTree>[0]);
  return printTree(tree);
}

let runner: BasicTestRunner;

beforeEach(async () => {
  runner = await createGraphDocsTestRunner();
});

describe('YamlFrontMatter', () => {
  it('renders YAML front matter block', () => {
    const result = renderToString(
      <YamlFrontMatter
        title='test resource type'
        description='A test description.'
        docType='resourcePageType'
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain('---');
    expect(result).toContain('title: "test resource type"');
    expect(result).toContain('description: "A test description."');
    expect(result).toContain('doc_type: resourcePageType');
    expect(result).toContain('ms.date: 01/15/2025');
  });

  it('escapes quotes in description', () => {
    const result = renderToString(
      <YamlFrontMatter
        title='test'
        description='A "quoted" description.'
        docType='apiPageType'
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain('description: "A \\"quoted\\" description."');
  });

  it('renders author field when provided', () => {
    const result = renderToString(
      <YamlFrontMatter
        title='test'
        description='desc'
        docType='resourcePageType'
        msDate='01/15/2025'
        author='jasonjoh'
      />,
    );

    expect(result).toContain('author: jasonjoh');
  });

  it('defaults author field when not provided', () => {
    const result = renderToString(
      <YamlFrontMatter
        title='test'
        description='desc'
        docType='resourcePageType'
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain('author: YOUR_GITHUB_USERNAME');
  });
});

describe('PropertiesTable', () => {
  it('renders property rows for entity model', async () => {
    await runner.compile(`
      using MsGraph;

      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        @entity model testEntity {
          @readOnly @computed @key id: string;
          /** The display name. */
          @computed displayName: string;
          @computed count: int32;
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <PropertiesTable program={runner.program} model={entity.model} />,
    );

    expect(result).toContain('## Properties');
    expect(result).toContain('| Property | Type | Description |');
    expect(result).toContain('| id | String |');
    expect(result).toContain('| displayName | String | The display name. |');
    expect(result).toContain('| count | Int32 |');
  });

  it('excludes @contains properties', async () => {
    await runner.compile(`
      using MsGraph;

      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        @entity model child {
          @readOnly @computed @key id: string;
        }

        @entity model parent {
          @readOnly @computed @key id: string;
          @contains children: child[];
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    const parentEntity = types.entities.find((e) => e.name === 'parent')!;

    const result = renderToString(
      <PropertiesTable program={runner.program} model={parentEntity.model} />,
    );

    expect(result).not.toContain('children');
    expect(result).toContain('| id |');
  });
});

describe('RelationshipsTable', () => {
  it('renders @contains navigation properties', async () => {
    await runner.compile(`
      using MsGraph;

      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        @entity model child {
          @readOnly @computed @key id: string;
        }

        @entity model parent {
          @readOnly @computed @key id: string;
          /** The children items. */
          @contains children: child[];
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    const parentEntity = types.entities.find((e) => e.name === 'parent')!;

    const result = renderToString(
      <RelationshipsTable
        program={runner.program}
        model={parentEntity.model}
      />,
    );

    expect(result).toContain('## Relationships');
    expect(result).toContain('| Relationship | Type | Description |');
    expect(result).toContain('| children |');
    expect(result).toContain('The children items.');
  });

  it('returns empty when no containment properties', async () => {
    await runner.compile(`
      using MsGraph;

      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        @entity model simple {
          @readOnly @computed @key id: string;
          @computed name: string;
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'simple')!;

    const result = renderToString(
      <RelationshipsTable program={runner.program} model={entity.model} />,
    );

    expect(result).not.toContain('## Relationships');
  });
});

describe('EnumMembersTable', () => {
  it('renders enum members with values', async () => {
    await runner.compile(`
      using MsGraph;

      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        enum testStatus {
          /** The item is active. */
          active,
          /** The item is inactive. */
          inactive,
          unknownFutureValue
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    const enumInfo = types.enums.find((e) => e.name === 'testStatus')!;

    const result = renderToString(
      <EnumMembersTable
        program={runner.program}
        enumType={enumInfo.enumType}
      />,
    );

    expect(result).toContain('| Member | Value | Description |');
    expect(result).toContain('| active |');
    expect(result).toContain('The item is active.');
    expect(result).toContain('| inactive |');
    expect(result).toContain('| unknownFutureValue |');
  });
});

describe('JsonRepresentation', () => {
  it('renders JSON block with @odata.type and properties', async () => {
    await runner.compile(`
      using MsGraph;

      @publicNamespace("microsoft.graph")
      namespace microsoft.graph {
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @computed displayName: string;
          @computed count: int32;
        }
      }
    `);

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <JsonRepresentation program={runner.program} model={entity.model} />,
    );

    expect(result).toContain('## JSON representation');
    expect(result).toContain('"@odata.type": "#microsoft.graph.testEntity"');
    expect(result).toContain('"id": "String"');
    expect(result).toContain('"displayName": "String"');
    expect(result).toContain('"count": 0');
  });
});

describe('MethodsTable', () => {
  it('renders method rows with links', () => {
    const ops = [
      {
        name: 'get testItem',
        httpMethod: 'GET',
        routePath: 'items/{id}',
        resourceTypeName: 'items',
        description: 'Get a test item.',
        docKind: DocOperationKind.GetResource,
        returnTypeName: 'testItem',
      },
    ];

    const result = renderToString(
      <MethodsTable
        operations={ops}
        getMethodFilename={() => 'testitem-get.md'}
      />,
    );

    expect(result).toContain('## Methods');
    expect(result).toContain('| Method | Return Type | Description |');
    expect(result).toContain('[get testItem](../api/testitem-get.md)');
    expect(result).toContain('Get a test item.');
  });

  it('returns empty when no operations', () => {
    const result = renderToString(
      <MethodsTable operations={[]} getMethodFilename={() => ''} />,
    );

    expect(result).not.toContain('## Methods');
  });
});
