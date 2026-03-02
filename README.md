# @microsoft/typespec-graph-docs-emitter

<!-- cSpell:ignore typespec msgraph tspconfig -->
<!-- cSpell:ignoreRegexp copilotconversation.* -->

A TypeSpec emitter that generates Markdown API reference documentation for Microsoft Graph APIs, matching the format used on [learn.microsoft.com](https://learn.microsoft.com/graph/api/resources/).

## Features

- Generates **resource type pages** with properties, relationships, methods, and JSON representation
- Generates **API method pages** with HTTP request, headers, request body, and response sections
- Generates **enum type pages** and **complex type pages**
- Reads `@microsoft/typespec-msgraph` decorators (`@entity`, `@complex`, `@graphRoute`, `@contains`, etc.)
- Produces YAML front matter compatible with the Microsoft Learn publishing pipeline
- Follows learn.microsoft.com filename conventions

## Installation

```sh
npm install @microsoft/typespec-graph-docs-emitter
```

### Peer Dependencies

This emitter requires the following peer dependencies:

- `@typespec/compiler` ^1.9.0
- `@alloy-js/core` ^0.22.0
- `@microsoft/typespec-msgraph` >=1.0.0

## Usage

Add the emitter to your `tspconfig.yaml`:

```yaml
emit:
  - "@microsoft/typespec-graph-docs-emitter"
options:
  "@microsoft/typespec-graph-docs-emitter":
    api-version: "beta"           # optional, default: "v1.0"
    output-dir: "./docs"          # optional, overrides compiler default
    ms-date: "2025-01-15"         # optional, overrides ms.date in YAML front matter
```

Then run the TypeSpec compiler:

```sh
npx tsp compile .
```

## Options

| Option        | Type     | Default          | Description                                                                                 |
|---------------|----------|------------------|---------------------------------------------------------------------------------------------|
| `api-version` | `string` | `v1.0`           | The Microsoft Graph API version (e.g., `v1.0`, `beta`). Used for display in generated docs. |
| `output-dir`  | `string` | compiler default | Override the output directory for generated files.                                          |
| `ms-date`     | `string` | today's date     | Override the `ms.date` value in YAML front matter (ISO 8601 date).                          |

## Output Structure

```text
{output-dir}/
├── resources/
│   ├── copilotconversation.md          # Entity type pages
│   ├── copilotconversationlocation.md  # Complex type pages
│   └── copilotconversationstate.md     # Enum type pages
└── api/
    ├── copilotconversation-get.md
    ├── copilot-list-conversations.md
    ├── copilot-post-conversations.md
    └── copilotconversation-chat.md
```

### Filename Conventions

**Resource/enum/complex types:** `{resourcename}.md` — all lowercase, no hyphens.

**API methods:**

| Operation        | Filename Pattern                | Example                  |
|------------------|---------------------------------|--------------------------|
| GET (collection) | `{parent}-list-{collection}.md` | `user-list-calendars.md` |
| GET (single)     | `{resource}-get.md`             | `message-get.md`         |
| POST (create)    | `{parent}-post-{collection}.md` | `user-post-messages.md`  |
| POST (action)    | `{resource}-{action}.md`        | `message-reply.md`       |
| PATCH            | `{resource}-update.md`          | `message-update.md`      |
| DELETE           | `{resource}-delete.md`          | `message-delete.md`      |

## Development

```sh
npm run build          # Build with alloy
npm test               # Run tests (vitest)
npm run lint           # ESLint + Prettier
npm run clean          # Remove dist/
```

## License

MIT
