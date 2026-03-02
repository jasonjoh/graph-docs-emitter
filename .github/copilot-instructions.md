# Copilot Instructions for graph-docs-emitter

## Project Overview

This is `@microsoft/typespec-graph-docs-emitter`, a TypeSpec emitter that generates Markdown API reference documentation for Microsoft Graph APIs. It reads TypeSpec definitions annotated with `@microsoft/typespec-msgraph` decorators and produces Markdown files matching the format used on learn.microsoft.com.

The emitter uses the **alloy-based emitter framework** (`@typespec/emitter-framework`) with JSX components for Markdown generation. The `.tsp` files in `sample-specs/` are test fixtures, not the primary artifacts.

## Architecture

```
src/
├── index.tsx                 # $onEmit entry point, emitter orchestration
├── lib.ts                    # createTypeSpecLibrary() definition, options schema
├── utils/
│   ├── type-collector.ts     # Walks program, collects Graph entities/complex/enums/routes
│   ├── graph-metadata.ts     # Reads Graph-specific decorators (@graphRoute, @entity, etc.)
│   ├── operation-resolver.ts # Maps GraphOps.* templates → HTTP verbs + params
│   ├── type-formatter.ts     # Formats TypeSpec types as display strings with MD links
│   └── filename.ts           # Generates output filenames per naming conventions
├── components/
│   ├── YamlFrontMatter.tsx   # YAML front matter block
│   ├── PropertiesTable.tsx   # Property | Type | Description table
│   ├── MethodsTable.tsx      # Methods section with links to method pages
│   ├── RelationshipsTable.tsx# @contains navigation properties table
│   ├── JsonRepresentation.tsx# Sample JSON from model properties
│   ├── EnumMembersTable.tsx  # Enum Member | Value | Description table
│   ├── ResourceTypePage.tsx  # Full resource type page
│   ├── ApiMethodPage.tsx     # Full API method page
│   ├── EnumTypePage.tsx      # Full enum type page
│   └── ComplexTypePage.tsx   # Full complex type page
test/
├── test-host.ts              # Test infrastructure (TypeSpec test host + runner)
├── utils/
│   ├── filename.test.ts      # Unit tests for filename generation
│   └── components.test.tsx   # Component rendering tests
├── integration/
│   └── type-discovery.test.ts# Integration tests for type collection + operation resolution
sample-specs/                 # TypeSpec fixtures for testing
```

### Emitter Flow

1. `$onEmit` receives the compiled TypeSpec program via `EmitContext`
2. `type-collector` walks the program to discover Graph-annotated types
3. `graph-metadata` and `operation-resolver` extract decorator data and resolve operations
4. Page-level JSX components (`ResourceTypePage`, `ApiMethodPage`, etc.) compose section components
5. Alloy's `writeOutput` renders all `<SourceFile>` elements to disk

### Output Structure

```
{output-dir}/
├── resources/                # Resource type, complex type, and enum pages
│   ├── {resourcename}.md
│   ├── {complextype}.md
│   └── {enumname}.md
└── api/                      # API method pages
    ├── {resource}-get.md
    ├── {parent}-list-{collection}.md
    └── ...
```

### Key Dependencies

- `@typespec/compiler` — TypeSpec compilation, type system, `navigateProgram`
- `@typespec/emitter-framework` — `writeOutput`, rendering pipeline
- `@alloy-js/core` — `Output`, `SourceFile`, `SourceDirectory`, JSX rendering
- `@microsoft/typespec-msgraph` — Graph-specific decorators and shared models (peer dep)

### Key Design Decisions

- **`@microsoft/typespec-msgraph` exports limitation**: Only decorator accessor functions are publicly exported (e.g., `isEntity`, `isComplex`, `getGraphRoutePaths`). Entity types (`MsGraphModel`, etc.) and utility functions (`GetMsGraphNamespaces`, etc.) are internal. The emitter uses only public APIs + compiler's `navigateProgram`.
- **Operation resolution**: Prefers `sourceOperation.name` (the GraphOps template name) over the local alias name for pattern matching. This correctly distinguishes `get is GraphOps.GetResource` from `get is GraphOps.GetPagedCollection`.
- **JSX preserve mode**: `tsconfig.json` uses `jsx: 'preserve'` and the alloy rollup plugin handles the JSX transform. The `@alloy-js/core` package does not export a standard `jsx-dev-runtime`, so esbuild's automatic mode cannot be used. The resulting vitest warning about JSX import source is harmless.

## Emitter Options

Configured in `tspconfig.yaml` under the emitter name:

| Option | Type | Default | Description |
|---|---|---|---|
| `api-version` | string | `v1.0` | API version for display (e.g., `v1.0`, `beta`) |
| `output-dir` | string | compiler default | Override the output directory |
| `ms-date` | string | today's date | Override `ms.date` in YAML front matter (ISO 8601) |

## Output File Naming Conventions

### Resource, enum, and complex type pages

`{resourcename}.md` — all-lowercase, no spaces or hyphens. Example: `copilotconversation.md`.

### API method pages

| Operation | Pattern | Example |
|---|---|---|
| GET collection | `{parent-entity}-list-{plural-resource-name}.md` | `user-list-calendars.md` |
| GET single | `{resource-name}-get.md` | `message-get.md` |
| POST (creates item) | `{parent-entity}-post-{plural-resource-name}.md` | `user-post-messages.md` |
| POST (action/function) | `{resource-name}-{action-or-function}.md` | `message-reply.md` |
| PATCH | `{resource-name}-update.md` | `message-update.md` |
| DELETE | `{resource-name}-delete.md` | `message-delete.md` |

The parent entity is resolved from the route hierarchy. The resource name comes from the entity model name.

## Output Page Format (learn.microsoft.com)

### Resource type page

- YAML front matter: `title`, `description`, `doc_type: resourcePageType`, `ms.date`
- `# {name} resource type` heading, then `Namespace: microsoft.graph`
- Description, `## Methods` table, `## Properties` table, `## Relationships` table, `## JSON representation`

### API method page

- YAML front matter: `title`, `description`, `doc_type: apiPageType`, `ms.date`
- `# {title}` heading, then `Namespace: microsoft.graph`
- Description, `## HTTP request`, `## Request headers`, `## Request body`, `## Response`

### Enum type page

- YAML front matter, `# {name} enum type` heading, `## Members` table

### Complex type page

- Same as resource type page but without `## Methods` and `## Relationships` sections

## TypeSpec Input Conventions

### Graph Decorators the Emitter Must Read

- `@entity` / `@complex` — Classify models as OData entities vs complex types
- `@graphRoute("path")` — Route path for interfaces
- `@contains` — Navigation properties (emitted as Relationships)
- `@ownerless` — Singleton entry points
- `@computed`, `@readOnly`, `@requiredForCreate`, `@immutable`, `@key` — Property metadata
- `@publicNamespace("microsoft.graph")` — Namespace filtering (only emit types from these)
- `@operationParameters` — Identifies request body models (excluded from type pages)

### Operation Mapping

`GraphOps.*` templates map to HTTP verbs. The resolver prefers the source operation template name over the local alias:

| GraphOps Template | HTTP Method | Page Type |
|---|---|---|
| `GetCollection` / `GetPagedCollection` | GET | list |
| `GetResource` | GET | get |
| `Post` | POST | post (creates) |
| `Action<...>` | POST | action |
| `Function<...>` | GET | function |
| `PatchNoResponse` / `PatchWithResponse` | PATCH | update |
| `Put` / `PutWithResponse` | PUT | update |
| `Delete` | DELETE | delete |

### LocalOps Pattern

Workloads define a `LocalOps` interface in the `MsGraph` namespace with workload-wide defaults for standard operations. Route interfaces reference `LocalOps.*` instead of `GraphOps.*` directly.

### Enums

Always include `unknownFutureValue` as the last member (Microsoft Graph evolvable enum pattern).

## Code Style

Enforced by ESLint + Prettier (run `npm run lint`):

- **Single quotes** for all strings (including JSX attributes): `'foo'` not `"foo"`
- **Copyright header** required at the top of every `.ts`, `.tsx`, and `.js` file:
  ```ts
  // Copyright (c) Microsoft Corporation.
  // Licensed under the MIT license.

  ```
- **No unused variables** — prefix intentionally unused parameters with `_` (e.g., `_context`)
- **Prettier formatting**: 80-char print width, trailing commas, `endOfLine: 'auto'`
- 2-space indentation (spaces, not tabs), insert final newline
- Emitter source is TypeScript + JSX (`.tsx` for components, `.ts` for utilities)
- JSX files require `/** @jsxImportSource @alloy-js/core */` pragma
- camelCase for TypeSpec model/property/enum names in the sample specs

## Build, Test & Lint Commands

```sh
npm run build          # TypeScript + alloy build (npx alloy build)
npm test               # Run full test suite (vitest)
npx vitest run <file>  # Run a single test file
npm run lint           # ESLint (includes Prettier checks)
npx eslint --fix .     # Auto-fix lint issues
npm run clean          # Remove dist/
```

## Testing

Tests use `vitest` with the `@alloy-js/rollup-plugin` for JSX support.

- **Unit tests** (`test/utils/`): Pure function tests for filename generation and component rendering
- **Integration tests** (`test/integration/`): Use `@typespec/compiler/testing` and `@microsoft/typespec-msgraph/testing` to compile inline TypeSpec and verify type collection + operation resolution
- Test host creates a fresh `BasicTestRunner` per test via `beforeEach` to avoid state leakage
- JSX components are tested by rendering with `renderTree()` + `printTree()` from `@alloy-js/core`
