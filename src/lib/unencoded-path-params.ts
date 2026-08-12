import { z } from 'zod';

export type UnencodedPathParameterKind = 'quoted-literal' | 'nonnegative-integer' | 'relative-path';

const URL_DELIMITERS = /[?#\\]/;
const PERCENT_ESCAPE = /%[0-9a-f]{2}/i;

function containsControlCharacter(value: string): boolean {
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    if (code <= 0x1f || code === 0x7f) return true;
  }
  return false;
}

function placeholderBounds(pathPattern: string, paramName: string): [number, number] {
  const placeholder = `{${paramName}}`;
  const start = pathPattern.indexOf(placeholder);
  if (start < 0) {
    throw new Error(
      `skipEncoding parameter ${JSON.stringify(paramName)} is not present in ${JSON.stringify(pathPattern)}`
    );
  }
  return [start, start + placeholder.length];
}

export function classifyUnencodedPathParameter(
  pathPattern: string,
  paramName: string
): UnencodedPathParameterKind {
  const [start, end] = placeholderBounds(pathPattern, paramName);
  const before = pathPattern.slice(0, start);
  const after = pathPattern.slice(end);

  if (before.endsWith("='") && after.startsWith("')")) return 'quoted-literal';
  if (before.endsWith('index=') && after.startsWith(')')) return 'nonnegative-integer';
  if (before.endsWith(':/')) return 'relative-path';

  throw new Error(
    `skipEncoding parameter ${JSON.stringify(paramName)} has an unsupported route context in ${JSON.stringify(pathPattern)}`
  );
}

export function unencodedPathParameterError(
  pathPattern: string,
  paramName: string,
  value: unknown
): string | undefined {
  if (typeof value !== 'string') return `${paramName} must be a string`;
  if (value.length === 0) return `${paramName} must not be empty`;
  if (containsControlCharacter(value)) return `${paramName} must not contain control characters`;

  const kind = classifyUnencodedPathParameter(pathPattern, paramName);
  switch (kind) {
    case 'quoted-literal':
      // These values are interpolated inside an OData single-quoted path function.
      // Their content is encoded by prepareUnencodedPathParameter, so punctuation is
      // data rather than URL structure. Only controls remain invalid.
      return undefined;

    case 'nonnegative-integer':
      if (PERCENT_ESCAPE.test(value)) return `${paramName} must not contain percent-encoded bytes`;
      return /^(0|[1-9]\d*)$/.test(value)
        ? undefined
        : `${paramName} must be a nonnegative decimal integer`;

    case 'relative-path': {
      if (PERCENT_ESCAPE.test(value)) return `${paramName} must not contain percent-encoded bytes`;
      if (URL_DELIMITERS.test(value) || value.includes(':')) {
        return `${paramName} contains an unsafe route or URL delimiter`;
      }
      if (value.startsWith('/') || value.endsWith('/')) {
        return `${paramName} must be relative and must not start or end with a slash`;
      }
      const segments = value.split('/');
      if (segments.some((segment) => segment === '' || segment === '.' || segment === '..')) {
        return `${paramName} must contain nonempty relative path segments`;
      }
      return undefined;
    }
  }
}

export function assertSafeUnencodedPathParameter(
  pathPattern: string,
  paramName: string,
  value: unknown
): asserts value is string {
  const error = unencodedPathParameterError(pathPattern, paramName, value);
  if (error) throw new Error(`Unsafe unencoded path parameter: ${error}.`);
}

export function prepareUnencodedPathParameter(
  pathPattern: string,
  paramName: string,
  value: unknown
): string {
  assertSafeUnencodedPathParameter(pathPattern, paramName, value);
  return classifyUnencodedPathParameter(pathPattern, paramName) === 'quoted-literal'
    ? encodeURIComponent(value.replace(/'/g, "''")).replace(/'/g, '%27')
    : value;
}

export function refineUnencodedPathParameterSchema(
  schema: z.ZodTypeAny,
  pathPattern: string,
  paramName: string
): z.ZodTypeAny {
  return schema.superRefine((value, ctx) => {
    // Optional parameters remain optional. If supplied, runtime interpolation applies
    // the same validation again so execute-tool and stale clients cannot bypass it.
    if (value === undefined) return;
    const error = unencodedPathParameterError(pathPattern, paramName, value);
    if (error) ctx.addIssue({ code: z.ZodIssueCode.custom, message: error });
  });
}
