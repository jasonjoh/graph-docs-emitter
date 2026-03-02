// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/** @jsxImportSource @alloy-js/core */
import { Children } from '@alloy-js/core';

export interface YamlFrontMatterProps {
  title: string;
  description: string;
  docType: 'resourcePageType' | 'apiPageType' | 'enumPageType';
  msDate?: string;
  author?: string;
}

/**
 * Renders a YAML front matter block for a learn.microsoft.com doc page.
 */
export function YamlFrontMatter(props: YamlFrontMatterProps): Children {
  const date = props.msDate ?? new Date().toISOString().split('T')[0];
  return [
    '---\n',
    `title: "${props.title}"\n`,
    `description: "${escapeYaml(props.description)}"\n`,
    `author: ${props.author ? props.author : 'YOUR_GITHUB_USERNAME'}\n`,
    `doc_type: ${props.docType}\n`,
    `ms.date: ${date}\n`,
    '---\n',
  ];
}

function escapeYaml(value: string): string {
  return value.replace(/"/g, '\\"').replace(/\n/g, ' ');
}
