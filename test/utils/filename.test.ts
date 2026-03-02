// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, it, expect } from 'vitest';
import {
  getTypeFilename,
  getMethodFilename,
} from '../../src/utils/filename.js';
import {
  DocOperationKind,
  ResolvedOperation,
} from '../../src/utils/operation-resolver.js';

function makeOp(
  overrides: Partial<ResolvedOperation> & { docKind: DocOperationKind },
): ResolvedOperation {
  return {
    name: 'test',
    httpMethod: 'GET',
    routePath: 'test',
    resourceTypeName: 'test',
    description: undefined,
    returnTypeName: undefined,
    ...overrides,
  };
}

describe('getTypeFilename', () => {
  it('lowercases the type name', () => {
    expect(getTypeFilename('CopilotConversation')).toBe(
      'copilotconversation.md',
    );
  });

  it('handles already-lowercase names', () => {
    expect(getTypeFilename('message')).toBe('message.md');
  });

  it('strips nothing else from simple names', () => {
    expect(getTypeFilename('appQuotaSettings')).toBe('appquotasettings.md');
  });
});

describe('getMethodFilename', () => {
  it('generates GET collection filename: {parent}-list-{collection}', () => {
    const op = makeOp({
      docKind: DocOperationKind.ListCollection,
      parentEntityName: 'copilot',
      resourceTypeName: 'conversations',
    });
    expect(getMethodFilename(op, 'copilotConversation')).toBe(
      'copilot-list-conversations.md',
    );
  });

  it('falls back to entity name when no parent for collection', () => {
    const op = makeOp({
      docKind: DocOperationKind.ListCollection,
      parentEntityName: undefined,
      resourceTypeName: 'messages',
    });
    expect(getMethodFilename(op, 'Message')).toBe('message-list-messages.md');
  });

  it('generates GET single resource filename: {resource}-get', () => {
    const op = makeOp({
      docKind: DocOperationKind.GetResource,
    });
    expect(getMethodFilename(op, 'Message')).toBe('message-get.md');
  });

  it('generates POST create filename: {parent}-post-{collection}', () => {
    const op = makeOp({
      docKind: DocOperationKind.PostCreate,
      parentEntityName: 'copilot',
      resourceTypeName: 'conversations',
    });
    expect(getMethodFilename(op, 'copilotConversation')).toBe(
      'copilot-post-conversations.md',
    );
  });

  it('generates action filename: {resource}-{action}', () => {
    const op = makeOp({
      docKind: DocOperationKind.Action,
      actionOrFunctionName: 'chat',
    });
    expect(getMethodFilename(op, 'copilotConversation')).toBe(
      'copilotconversation-chat.md',
    );
  });

  it('generates function filename: {resource}-{function}', () => {
    const op = makeOp({
      docKind: DocOperationKind.Function,
      actionOrFunctionName: 'getApiUsage',
    });
    expect(getMethodFilename(op, 'reportRoot')).toBe(
      'reportroot-getapiusage.md',
    );
  });

  it('generates PATCH filename: {resource}-update', () => {
    const op = makeOp({
      docKind: DocOperationKind.Update,
    });
    expect(getMethodFilename(op, 'Message')).toBe('message-update.md');
  });

  it('generates DELETE filename: {resource}-delete', () => {
    const op = makeOp({
      docKind: DocOperationKind.Delete,
    });
    expect(getMethodFilename(op, 'Message')).toBe('message-delete.md');
  });
});
