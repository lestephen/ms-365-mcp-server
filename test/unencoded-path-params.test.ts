import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
import {
  classifyUnencodedPathParameter,
  prepareUnencodedPathParameter,
  unencodedPathParameterError,
} from '../src/lib/unencoded-path-params.js';

interface EndpointConfig {
  toolName: string;
  pathPattern: string;
  skipEncoding?: string[];
}

const endpointsPath = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
  'src',
  'endpoints.json'
);
const endpoints = JSON.parse(readFileSync(endpointsPath, 'utf8')) as EndpointConfig[];

describe('unencoded path parameter policies', () => {
  it('classifies every skipEncoding interpolation into a supported safe policy', () => {
    const inventory = endpoints.flatMap((endpoint) =>
      (endpoint.skipEncoding ?? []).map((paramName) => ({
        toolName: endpoint.toolName,
        paramName,
        kind: classifyUnencodedPathParameter(endpoint.pathPattern, paramName),
      }))
    );

    expect(inventory).toHaveLength(16);
    expect(inventory.filter((item) => item.paramName === 'q')).toEqual([
      {
        toolName: 'search-onedrive-files',
        paramName: 'q',
        kind: 'quoted-literal',
      },
    ]);
    expect(inventory.filter((item) => item.paramName === 'address')).toHaveLength(12);
    expect(
      inventory
        .filter((item) => item.paramName === 'address')
        .every((item) => item.kind === 'quoted-literal')
    ).toBe(true);
    expect(inventory.filter((item) => item.paramName === 'index')).toHaveLength(2);
    expect(
      inventory
        .filter((item) => item.paramName === 'index')
        .every((item) => item.kind === 'nonnegative-integer')
    ).toBe(true);
    expect(inventory.filter((item) => item.paramName === 'path')).toEqual([
      {
        toolName: 'get-sharepoint-site-by-path',
        paramName: 'path',
        kind: 'relative-path',
      },
    ]);
  });

  it('allows the syntax each raw interpolation actually needs', () => {
    expect(
      unencodedPathParameterError("/drives/{drive-id}/search(q='{q}')", 'q', 'quarterly report')
    ).toBeUndefined();
    expect(
      prepareUnencodedPathParameter("/drives/{drive-id}/search(q='{q}')", 'q', "Stephen's report")
    ).toBe("Stephen''s report");
    expect(
      unencodedPathParameterError(
        "/workbook/worksheets/{id}/range(address='{address}')",
        'address',
        '$A$1:Z99'
      )
    ).toBeUndefined();
    expect(
      unencodedPathParameterError('/rows/itemAt(index={index})', 'index', '12')
    ).toBeUndefined();
    expect(
      unencodedPathParameterError('/sites/{site-id}:/{path}', 'path', 'sites/marketing/north')
    ).toBeUndefined();
  });

  it('rejects breakout syntax for quoted literals, integer functions, and relative paths', () => {
    expect(
      unencodedPathParameterError("/drives/{drive-id}/search(q='{q}')", 'q', "report')/children")
    ).toMatch(/slash/);
    expect(
      unencodedPathParameterError(
        "/workbook/worksheets/{id}/range(address='{address}')",
        'address',
        'A1)/tables?x=1'
      )
    ).toMatch(/slash|URL delimiter/);
    expect(
      unencodedPathParameterError('/rows/itemAt(index={index})', 'index', '0)/tables')
    ).toMatch(/nonnegative decimal integer/);
    expect(
      unencodedPathParameterError(
        '/sites/{site-id}:/{path}',
        'path',
        'sites/marketing:/lists/blocked'
      )
    ).toMatch(/unsafe route/);
  });
});
