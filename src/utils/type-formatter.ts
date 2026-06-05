// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Type,
  Model,
  Scalar,
  Union,
  isArrayModelType,
  isRecordModelType,
} from '@typespec/compiler';
import { DEFAULT_NAMESPACE } from './graph-metadata.js';

/**
 * Format a TypeSpec type as a display string for documentation.
 * Produces Markdown links for model/enum references.
 */
export function formatTypeName(type: Type): string {
  switch (type.kind) {
    case 'Model':
      return formatModelType(type);
    case 'Scalar':
      return formatScalarType(type);
    case 'Enum':
      return `[${type.name}](${type.name.toLowerCase()}.md)`;
    case 'Union':
      return formatUnionType(type);
    case 'Intrinsic':
      if (type.name === 'void') return 'None';
      if (type.name === 'null') return 'null';
      if (type.name === 'unknown') return 'Json';
      return type.name;
    default:
      return formatLiteralType(type);
  }
}

function formatModelType(model: Model): string {
  // Array types (T[]) render as "ElementType collection"
  if (isArrayModelType(model)) {
    const elementType = model.indexer!.value;
    return `${formatTypeName(elementType)} collection`;
  }

  // Named models (including named maps like `model X is Record<T>`)
  // render as a linked reference
  if (model.name && model.name !== 'Record') {
    return `[${model.name}](${model.name.toLowerCase()}.md)`;
  }

  // Anonymous or unnamed record/models
  return 'Json';
}

function formatScalarType(scalar: Scalar): string {
  return SCALAR_TYPE_MAP[scalar.name] ?? scalar.name;
}

function formatUnionType(union: Union): string {
  const variants = [...union.variants.values()];

  // Filter out null for nullable types
  const nonNullVariants = variants.filter(
    (v) => !(v.type.kind === 'Intrinsic' && v.type.name === 'null'),
  );

  const isNullable = nonNullVariants.length < variants.length;

  if (nonNullVariants.length === 1) {
    return formatTypeName(nonNullVariants[0].type);
  }

  // Multiple non-null variants: join with " or ", append nullable indicator
  const formatted = nonNullVariants
    .map((v) => formatTypeName(v.type))
    .join(' or ');
  return isNullable ? `${formatted} (nullable)` : formatted;
}

function formatLiteralType(type: Type): string {
  if (type.kind === 'String') return 'String';
  if (type.kind === 'Number') return 'Int32';
  if (type.kind === 'Boolean') return 'Boolean';
  return 'String';
}

/**
 * Map TypeSpec scalar names to Microsoft Graph documentation type names.
 */
const SCALAR_TYPE_MAP: Record<string, string> = {
  string: 'String',
  boolean: 'Boolean',
  int8: 'Byte',
  int16: 'Int16',
  int32: 'Int32',
  int64: 'Int64',
  float32: 'Single',
  float64: 'Double',
  decimal: 'Decimal',
  decimal128: 'Decimal',
  bytes: 'Binary',
  plainDate: 'Date',
  plainTime: 'TimeOfDay',
  utcDateTime: 'DateTimeOffset',
  offsetDateTime: 'DateTimeOffset',
  duration: 'Duration',
  url: 'String',
  numeric: 'Double',
  integer: 'Int64',
  float: 'Double',
  safeint: 'Int64',
  uint8: 'Byte',
  uint16: 'Int32',
  uint32: 'Int64',
  uint64: 'Int64',
  // Graph-specific
  dateTimeOffset: 'DateTimeOffset',
  webUrl: 'String',
  stream: 'Stream',
};

/**
 * Format a TypeSpec type as a JSON representation value.
 * Used for generating the ## JSON representation section.
 */
export function formatJsonValue(type: Type): string {
  switch (type.kind) {
    case 'Scalar':
      return getJsonPlaceholder(type.name);
    case 'Model':
      if (isArrayModelType(type)) {
        return `[ ${formatJsonValue(type.indexer!.value)} ]`;
      }
      if (type.name && type.name !== 'Record') {
        return `{"@odata.type": "${DEFAULT_NAMESPACE}.${type.name}"}`;
      }
      if (!type.name || isRecordModelType(type)) return '{}';
      return `{"@odata.type": "${DEFAULT_NAMESPACE}.${type.name}"}`;
    case 'Enum':
      return '"String"';
    case 'Union':
      return formatJsonValueForUnion(type);
    case 'Intrinsic':
      if (type.name === 'null') return 'null';
      return '"String"';
    default:
      return '"String"';
  }
}

function formatJsonValueForUnion(union: Union): string {
  const variants = [...union.variants.values()];
  const nonNull = variants.filter(
    (v) => !(v.type.kind === 'Intrinsic' && v.type.name === 'null'),
  );
  if (nonNull.length === 1) {
    return formatJsonValue(nonNull[0].type);
  }
  return '"String"';
}

function getJsonPlaceholder(scalarName: string): string {
  switch (scalarName) {
    case 'string':
    case 'url':
      return '"String"';
    case 'boolean':
      return 'true';
    case 'int8':
    case 'int16':
    case 'int32':
    case 'int64':
    case 'integer':
    case 'safeint':
    case 'uint8':
    case 'uint16':
    case 'uint32':
    case 'uint64':
      return '0';
    case 'float32':
    case 'float64':
    case 'float':
    case 'decimal':
    case 'decimal128':
    case 'numeric':
      return '0.0';
    case 'plainDate':
      return '"Date"';
    case 'plainTime':
      return '"TimeOfDay"';
    case 'utcDateTime':
    case 'offsetDateTime':
    case 'dateTimeOffset':
      return '"String (timestamp)"';
    case 'duration':
      return '"Duration"';
    case 'bytes':
      return '"Binary"';
    default:
      return '"String"';
  }
}
