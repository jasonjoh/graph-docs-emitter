// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Enum, Program, getDoc } from '@typespec/compiler';
import { YamlFrontMatter } from './YamlFrontMatter.jsx';
import { GraphEnumInfo } from '../utils/type-collector.js';

export interface EnumsPageProps {
  program: Program;
  enums: GraphEnumInfo[];
  namespace: string;
  apiVersion?: string;
  msDate?: string;
  author?: string;
}

/**
 * Renders a Member | Value | Description table for an enum.
 */
function EnumMembersList(props: {
  enumType: Enum;
  program: Program;
}): Children {
  const rows: string[] = [];
  for (const [name, member] of props.enumType.members) {
    const value = member.value !== undefined ? String(member.value) : name;
    const description = getDoc(props.program, member) ?? '';
    rows.push(`| ${name} | ${value} | ${description} |`);
  }

  return [
    '\n| Member | Value | Description |\n',
    '|:--|:--|:--|\n',
    ...rows.map((r) => r + '\n'),
  ];
}

/**
 * Renders a combined enums page with all enum types.
 */
export function EnumsPage(props: EnumsPageProps): Children {
  const title = 'Enum types';
  const desc = 'Enum types for the Microsoft Graph API.';
  const isBeta = props.apiVersion === 'beta';

  const sorted = [...props.enums].sort((a, b) => a.name.localeCompare(b.name));

  return [
    <YamlFrontMatter
      title={title}
      description={desc}
      docType='enumPageType'
      msDate={props.msDate}
      author={props.author}
    />,
    '\n',
    `# ${title}\n\n`,
    `Namespace: ${props.namespace}\n\n`,
    isBeta
      ? '[!INCLUDE [beta-disclaimer](../../includes/beta-disclaimer.md)]\n\n'
      : '',
    ...sorted.flatMap((enumInfo) => [
      `### ${enumInfo.name} values\n`,
      <EnumMembersList enumType={enumInfo.enumType} program={props.program} />,
      '\n',
    ]),
  ];
}
