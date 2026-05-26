// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// cSpell:ignore jasonjoh testitem baseentity childentity

/** @jsxImportSource @alloy-js/core */
import { describe, it, expect, beforeEach } from 'vitest';
import { BasicTestRunner } from '@typespec/compiler/testing';
import { renderTree, printTree } from '@alloy-js/core';
import { createGraphDocsTestRunner, graphSpec } from '../test-host.js';
import { collectGraphTypes } from '../../src/utils/type-collector.js';
import {
  PropertiesTable,
  TODO_DESCRIPTION,
  hasMissingDescriptions,
} from '../../src/components/PropertiesTable.jsx';
import { RelationshipsTable } from '../../src/components/RelationshipsTable.jsx';
import { EnumsPage } from '../../src/components/EnumsPage.jsx';
import { JsonRepresentation } from '../../src/components/JsonRepresentation.jsx';
import { YamlFrontMatter } from '../../src/components/YamlFrontMatter.jsx';
import { MethodsTable } from '../../src/components/MethodsTable.jsx';
import { ResourceTypePage } from '../../src/components/ResourceTypePage.jsx';
import { ComplexTypePage } from '../../src/components/ComplexTypePage.jsx';
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
    expect(result).toContain('ms.topic: reference');
    expect(result).toContain('ms.localizationpriority: medium');
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

    expect(result).toContain('author: "jasonjoh"');
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

    expect(result).toContain('author: "YOUR_GITHUB_USERNAME"');
  });
});

describe('PropertiesTable', () => {
  it('renders property rows for entity model', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          /** The display name. */
          @computed displayName: string;
          @computed count: int32;
        }
    `),
    );

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
    await runner.compile(
      graphSpec(`
        @entity model child {
          @readOnly @computed @key id: string;
        }

        @entity model parent {
          @readOnly @computed @key id: string;
          /** The children items. */
          @contains children: child[];
        }
    `),
    );

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
    await runner.compile(
      graphSpec(`
        @entity model simple {
          @readOnly @computed @key id: string;
          @computed name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'simple')!;

    const result = renderToString(
      <RelationshipsTable program={runner.program} model={entity.model} />,
    );

    expect(result).toContain('## Relationships');
    expect(result).toContain('None.');
  });
});

describe('EnumsPage', () => {
  it('renders all enums in a single page', async () => {
    await runner.compile(
      graphSpec(`
        enum testStatus {
          /** The item is active. */
          active,
          /** The item is inactive. */
          inactive,
          unknownFutureValue
        }
    `),
    );

    const types = collectGraphTypes(runner.program);

    const result = renderToString(
      <EnumsPage
        program={runner.program}
        enums={types.enums}
        namespace='microsoft.graph'
      />,
    );

    expect(result).toContain('### testStatus values');
    expect(result).toContain('| Member |');
    expect(result).toContain('| active |');
    expect(result).toContain('| inactive |');
    expect(result).toContain('| unknownFutureValue |');
  });
});

describe('JsonRepresentation', () => {
  it('renders JSON block with @odata.type and properties', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @computed displayName: string;
          @computed count: int32;
        }
    `),
    );

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

describe('Placeholder descriptions', () => {
  it('shows TODO placeholder for properties without descriptions', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <PropertiesTable program={runner.program} model={entity.model} />,
    );

    expect(result).toContain(`| id | String | ${TODO_DESCRIPTION} |`);
    expect(result).toContain(`| name | String | ${TODO_DESCRIPTION} |`);
  });

  it('does not show TODO placeholder when descriptions are present', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          /** The unique ID. */
          @readOnly @computed @key id: string;
          /** The display name. */
          name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <PropertiesTable program={runner.program} model={entity.model} />,
    );

    expect(result).not.toContain(TODO_DESCRIPTION);
    expect(result).toContain('| id | String | The unique ID. |');
    expect(result).toContain('| name | String | The display name. |');
  });

  it('shows TODO placeholder for relationships without descriptions', async () => {
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
    const parentEntity = types.entities.find((e) => e.name === 'parent')!;

    const result = renderToString(
      <RelationshipsTable
        program={runner.program}
        model={parentEntity.model}
      />,
    );

    expect(result).toContain(TODO_DESCRIPTION);
    expect(result).toContain(`| children |`);
  });

  it('hasMissingDescriptions returns true when properties lack docs', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          /** Has a doc. */
          name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    expect(hasMissingDescriptions(runner.program, entity.model)).toBe(true);
  });

  it('hasMissingDescriptions returns false when all properties have docs', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          /** The unique ID. */
          @readOnly @computed @key id: string;
          /** The display name. */
          name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    expect(hasMissingDescriptions(runner.program, entity.model)).toBe(false);
  });

  it('ResourceTypePage includes HTML comment when descriptions are missing', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <ResourceTypePage
        program={runner.program}
        model={entity.model}
        description='A test entity.'
        namespace='microsoft.graph'
        operations={[]}
        getMethodFilename={() => ''}
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain(
      '<!-- This file contains placeholder descriptions',
    );
    expect(result).toContain(
      'Please update the TypeSpec source with the missing descriptions',
    );
  });

  it('ResourceTypePage omits HTML comment when all descriptions present', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          /** The unique ID. */
          @readOnly @computed @key id: string;
          /** The display name. */
          name: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <ResourceTypePage
        program={runner.program}
        model={entity.model}
        description='A test entity.'
        namespace='microsoft.graph'
        operations={[]}
        getMethodFilename={() => ''}
        msDate='01/15/2025'
      />,
    );

    expect(result).not.toContain(
      '<!-- This file contains placeholder descriptions',
    );
  });

  it('ComplexTypePage includes HTML comment when descriptions are missing', async () => {
    await runner.compile(
      graphSpec(`
        @complex model testComplex {
          name: string;
          value: int32;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const complex = types.complexTypes.find((c) => c.name === 'testComplex')!;

    const result = renderToString(
      <ComplexTypePage
        program={runner.program}
        model={complex.model}
        description='A test complex type.'
        namespace='microsoft.graph'
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain(
      '<!-- This file contains placeholder descriptions',
    );
  });

  it('ComplexTypePage omits HTML comment when all descriptions present', async () => {
    await runner.compile(
      graphSpec(`
        @complex model testComplex {
          /** The name. */
          name: string;
          /** The value. */
          value: int32;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const complex = types.complexTypes.find((c) => c.name === 'testComplex')!;

    const result = renderToString(
      <ComplexTypePage
        program={runner.program}
        model={complex.model}
        description='A test complex type.'
        namespace='microsoft.graph'
        msDate='01/15/2025'
      />,
    );

    expect(result).not.toContain(
      '<!-- This file contains placeholder descriptions',
    );
  });
});

describe('PropertiesTable - base model inheritance', () => {
  it('includes properties inherited from base model', async () => {
    await runner.compile(
      graphSpec(`
        @entity model baseEntity {
          /** The unique ID. */
          @readOnly @computed @key id: string;
          /** Base name. */
          @computed baseProp: string;
        }
        @entity model childEntity extends baseEntity {
          /** Child value. */
          @computed childProp: int32;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const child = types.entities.find((e) => e.name === 'childEntity')!;

    const result = renderToString(
      <PropertiesTable program={runner.program} model={child.model} />,
    );

    expect(result).toContain('| childProp | Int32 | Child value. |');
    expect(result).toContain('| baseProp | String | Base name. |');
    expect(result).toContain('| id | String | The unique ID. |');
  });

  it('excludes @contains from inherited properties', async () => {
    await runner.compile(
      graphSpec(`
        @entity model item {
          @readOnly @computed @key id: string;
        }
        @entity model baseEntity {
          @readOnly @computed @key id: string;
          @contains items: item[];
          /** A value. */
          @computed baseProp: string;
        }
        @entity model childEntity extends baseEntity {
          /** Child value. */
          @computed childProp: int32;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const child = types.entities.find((e) => e.name === 'childEntity')!;

    const result = renderToString(
      <PropertiesTable program={runner.program} model={child.model} />,
    );

    expect(result).not.toContain('items');
    expect(result).toContain('| baseProp |');
    expect(result).toContain('| childProp |');
  });

  it('sorts all properties (own + inherited) alphabetically', async () => {
    await runner.compile(
      graphSpec(`
        @entity model baseEntity {
          /** An ID. */
          @readOnly @computed @key id: string;
          /** Zebra prop. */
          @computed zebra: string;
        }
        @entity model childEntity extends baseEntity {
          /** Alpha prop. */
          @computed alpha: string;
          /** Middle prop. */
          @computed middle: int32;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const child = types.entities.find((e) => e.name === 'childEntity')!;

    const result = renderToString(
      <PropertiesTable program={runner.program} model={child.model} />,
    );

    const lines = result
      .split('\n')
      .filter(
        (l: string) =>
          l.startsWith('| alpha') ||
          l.startsWith('| id') ||
          l.startsWith('| middle') ||
          l.startsWith('| zebra'),
      );
    expect(lines).toHaveLength(4);
    expect(lines[0]).toContain('| alpha');
    expect(lines[1]).toContain('| id');
    expect(lines[2]).toContain('| middle');
    expect(lines[3]).toContain('| zebra');
  });
});

describe('RelationshipsTable - base model inheritance', () => {
  it('includes @contains properties from base model', async () => {
    await runner.compile(
      graphSpec(`
        @entity model item {
          @readOnly @computed @key id: string;
        }
        @entity model thing {
          @readOnly @computed @key id: string;
        }
        @entity model baseEntity {
          @readOnly @computed @key id: string;
          /** Base items. */
          @contains items: item[];
        }
        @entity model childEntity extends baseEntity {
          /** Child things. */
          @contains things: thing[];
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const child = types.entities.find((e) => e.name === 'childEntity')!;

    const result = renderToString(
      <RelationshipsTable program={runner.program} model={child.model} />,
    );

    expect(result).toContain('| items |');
    expect(result).toContain('Base items.');
    expect(result).toContain('| things |');
    expect(result).toContain('Child things.');
  });
});

describe('Beta disclaimer', () => {
  it('ResourceTypePage renders beta disclaimer when apiVersion is beta', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          /** The ID. */
          @readOnly @computed @key id: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <ResourceTypePage
        program={runner.program}
        model={entity.model}
        description='A test entity.'
        namespace='microsoft.graph'
        operations={[]}
        getMethodFilename={() => ''}
        apiVersion='beta'
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain('[!INCLUDE [beta-disclaimer');
  });

  it('ResourceTypePage omits beta disclaimer when apiVersion is v1.0', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          /** The ID. */
          @readOnly @computed @key id: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <ResourceTypePage
        program={runner.program}
        model={entity.model}
        description='A test entity.'
        namespace='microsoft.graph'
        operations={[]}
        getMethodFilename={() => ''}
        apiVersion='v1.0'
        msDate='01/15/2025'
      />,
    );

    expect(result).not.toContain('[!INCLUDE [beta-disclaimer');
  });

  it('ComplexTypePage renders beta disclaimer when apiVersion is beta', async () => {
    await runner.compile(
      graphSpec(`
        @complex model testComplex {
          /** A value. */
          value: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const complex = types.complexTypes.find((c) => c.name === 'testComplex')!;

    const result = renderToString(
      <ComplexTypePage
        program={runner.program}
        model={complex.model}
        description='A complex type.'
        namespace='microsoft.graph'
        apiVersion='beta'
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain('[!INCLUDE [beta-disclaimer');
  });

  it('EnumsPage renders beta disclaimer when apiVersion is beta', async () => {
    await runner.compile(
      graphSpec(`
        enum testEnum { a, unknownFutureValue }
    `),
    );

    const types = collectGraphTypes(runner.program);

    const result = renderToString(
      <EnumsPage
        program={runner.program}
        enums={types.enums}
        namespace='microsoft.graph'
        apiVersion='beta'
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain('[!INCLUDE [beta-disclaimer');
  });

  it('EnumsPage omits beta disclaimer when apiVersion is not beta', async () => {
    await runner.compile(
      graphSpec(`
        enum testEnum { a, unknownFutureValue }
    `),
    );

    const types = collectGraphTypes(runner.program);

    const result = renderToString(
      <EnumsPage
        program={runner.program}
        enums={types.enums}
        namespace='microsoft.graph'
        apiVersion='v1.0'
        msDate='01/15/2025'
      />,
    );

    expect(result).not.toContain('[!INCLUDE [beta-disclaimer');
  });
});

describe('EnumsPage - sorting', () => {
  it('sorts enums alphabetically', async () => {
    await runner.compile(
      graphSpec(`
        enum zebraEnum { z, unknownFutureValue }
        enum alphaEnum { a, unknownFutureValue }
        enum middleEnum { m, unknownFutureValue }
    `),
    );

    const types = collectGraphTypes(runner.program);

    const result = renderToString(
      <EnumsPage
        program={runner.program}
        enums={types.enums}
        namespace='microsoft.graph'
      />,
    );

    const alphaIdx = result.indexOf('### alphaEnum values');
    const middleIdx = result.indexOf('### middleEnum values');
    const zebraIdx = result.indexOf('### zebraEnum values');

    expect(alphaIdx).toBeLessThan(middleIdx);
    expect(middleIdx).toBeLessThan(zebraIdx);
  });
});

describe('JsonRepresentation - edge cases', () => {
  it('renders collection/array property as empty array', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @computed tags: string[];
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <JsonRepresentation program={runner.program} model={entity.model} />,
    );

    expect(result).toContain('"tags": []');
  });

  it('renders enum property as "String"', async () => {
    await runner.compile(
      graphSpec(`
        enum testStatus { active, inactive, unknownFutureValue }
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @computed status: testStatus;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <JsonRepresentation program={runner.program} model={entity.model} />,
    );

    expect(result).toContain('"status": "String"');
  });

  it('renders nullable property using underlying type', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
          name: string | null;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <JsonRepresentation program={runner.program} model={entity.model} />,
    );

    expect(result).toContain('"name": "String"');
  });

  it('renders nested model as @odata.type reference', async () => {
    await runner.compile(
      graphSpec(`
        @complex model nested {
          value: string;
        }
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @computed detail: nested;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <JsonRepresentation program={runner.program} model={entity.model} />,
    );

    expect(result).toContain(
      '"detail": {"@odata.type": "microsoft.graph.nested"}',
    );
  });

  it('excludes @contains properties from JSON representation', async () => {
    await runner.compile(
      graphSpec(`
        @entity model child {
          @readOnly @computed @key id: string;
        }
        @entity model testEntity {
          @readOnly @computed @key id: string;
          @computed name: string;
          @contains children: child[];
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <JsonRepresentation program={runner.program} model={entity.model} />,
    );

    expect(result).not.toContain('children');
    expect(result).toContain('"name": "String"');
  });

  it('uses custom namespace in @odata.type', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          @readOnly @computed @key id: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <JsonRepresentation
        program={runner.program}
        model={entity.model}
        namespace='microsoft.graph.beta'
      />,
    );

    expect(result).toContain(
      '"@odata.type": "#microsoft.graph.beta.testEntity"',
    );
  });
});

describe('ResourceTypePage - description fallback', () => {
  it('uses default description when none provided', async () => {
    await runner.compile(
      graphSpec(`
        @entity model testEntity {
          /** The ID. */
          @readOnly @computed @key id: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const entity = types.entities.find((e) => e.name === 'testEntity')!;

    const result = renderToString(
      <ResourceTypePage
        program={runner.program}
        model={entity.model}
        description={undefined}
        namespace='microsoft.graph'
        operations={[]}
        getMethodFilename={() => ''}
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain('Represents a testEntity.');
  });
});

describe('ComplexTypePage - description fallback', () => {
  it('uses default description when none provided', async () => {
    await runner.compile(
      graphSpec(`
        @complex model testComplex {
          /** A prop. */
          value: string;
        }
    `),
    );

    const types = collectGraphTypes(runner.program);
    const complex = types.complexTypes.find((c) => c.name === 'testComplex')!;

    const result = renderToString(
      <ComplexTypePage
        program={runner.program}
        model={complex.model}
        description={undefined}
        namespace='microsoft.graph'
        msDate='01/15/2025'
      />,
    );

    expect(result).toContain('Represents a testComplex.');
  });
});
