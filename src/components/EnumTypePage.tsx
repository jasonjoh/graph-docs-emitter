// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';
import { Enum, Program } from '@typespec/compiler';
import { YamlFrontMatter } from './YamlFrontMatter.jsx';
import { EnumMembersTable } from './EnumMembersTable.jsx';

export interface EnumTypePageProps {
  program: Program;
  enumType: Enum;
  description: string | undefined;
  namespace: string;
}

/**
 * Renders a full enum type page matching the learn.microsoft.com format.
 */
export function EnumTypePage(props: EnumTypePageProps): Children {
  const title = `${props.enumType.name} enum type`;
  const desc = props.description ?? `Represents ${props.enumType.name}.`;

  return [
    <YamlFrontMatter title={title} description={desc} docType='enumPageType' />,
    '\n',
    `# ${title}\n\n`,
    `Namespace: ${props.namespace}\n\n`,
    desc + '\n',
    '\n## Members\n',
    <EnumMembersTable program={props.program} enumType={props.enumType} />,
  ];
}
