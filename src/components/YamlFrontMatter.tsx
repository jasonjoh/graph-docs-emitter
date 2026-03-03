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
  const today = new Date();
  const defaultDate = `${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}/${today.getFullYear()}`;
  const date = props.msDate ?? defaultDate;
  return [
    '---\n',
    `title: "${props.title}"\n`,
    `description: "${escapeYaml(props.description)}"\n`,
    `author: ${props.author ? props.author : 'YOUR_GITHUB_USERNAME'}\n`,
    'ms.topic: reference\n',
    `ms.date: ${date}\n`,
    'ms.localizationpriority: medium\n',
    `doc_type: ${props.docType}\n`,
    '---\n',
  ];
}

function escapeYaml(value: string): string {
  return value.replace(/"/g, '\\"').replace(/\n/g, ' ');
}
