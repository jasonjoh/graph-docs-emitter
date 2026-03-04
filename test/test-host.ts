// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createTestHost, createTestRunner } from '@typespec/compiler/testing';
import { MsGraphTestLibrary } from '@microsoft/typespec-msgraph/testing';
import { HttpTestLibrary } from '@typespec/http/testing';

import type { BasicTestRunner, TestHost } from '@typespec/compiler/testing';

export async function createGraphDocsTestHost(): Promise<TestHost> {
  return createTestHost({
    libraries: [MsGraphTestLibrary, HttpTestLibrary],
  });
}

export async function createGraphDocsTestRunner(): Promise<BasicTestRunner> {
  const host = await createGraphDocsTestHost();
  return createTestRunner(host);
}

/**
 * Wraps TypeSpec body code in the standard Graph namespace boilerplate:
 * `using MsGraph; @publicNamespace("microsoft.graph") namespace microsoft.graph { ... }`
 */
export function graphSpec(body: string): string {
  return `
using MsGraph;

@publicNamespace("microsoft.graph")
namespace microsoft.graph {
${body}
}
`;
}
