// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Enum } from '@typespec/compiler';
import { getDoc } from '@typespec/compiler';
import { Program } from '@typespec/compiler';

export interface EnumMembersTableProps {
  program: Program;
  enumType: Enum;
}

/**
 * Renders a Markdown table of enum members with Member, Value, and Description columns.
 */
export function EnumMembersTable(props: EnumMembersTableProps): Children {
  const rows: string[] = [];

  for (const [name, member] of props.enumType.members) {
    const description = getDoc(props.program, member) ?? '';
    const value =
      member.value !== undefined ? String(member.value) : String(name);
    rows.push(`| ${name} | \`${value}\` | ${description} |`);
  }

  return [
    '\n| Member | Value | Description |\n',
    '|:---|:---|:---|\n',
    ...rows.map((r) => r + '\n'),
  ];
}
