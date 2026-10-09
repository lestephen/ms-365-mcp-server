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

    expect(inventory).toHaveLength(18);
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
    expect(inventory.filter((item) => item.kind === 'datetime-literal')).toEqual([
      {
        toolName: 'list-adhoc-call-transcripts',
        paramName: 'startDateTime',
        kind: 'datetime-literal',
      },
      {
        toolName: 'list-adhoc-call-transcripts',
        paramName: 'endDateTime',
        kind: 'datetime-literal',
      },
    ]);
  });

  it('encodes quoted literal content while preserving supported function syntax', () => {
    expect(
      unencodedPathParameterError("/drives/{drive-id}/search(q='{q}')", 'q', 'Budget#1/? 100%')
    ).toBeUndefined();
    expect(
      prepareUnencodedPathParameter("/drives/{drive-id}/search(q='{q}')", 'q', "Stephen's report")
    ).toBe('Stephen%27%27s%20report');
    expect(
      prepareUnencodedPathParameter("/drives/{drive-id}/search(q='{q}')", 'q', 'Budget#1/? 100%')
    ).toBe('Budget%231%2F%3F%20100%25');
    expect(
      prepareUnencodedPathParameter(
        "/workbook/worksheets/{id}/range(address='{address}')",
        'address',
        '$A$1:Z99'
      )
    ).toBe('%24A%241%3AZ99');
    expect(
      prepareUnencodedPathParameter(
        "/workbook/worksheets/{id}/range(address='{address}')",
        'address',
        'Table1[#All]'
      )
    ).toBe('Table1%5B%23All%5D');
    expect(
      unencodedPathParameterError('/rows/itemAt(index={index})', 'index', '12')
    ).toBeUndefined();
    expect(
      unencodedPathParameterError('/sites/{site-id}:/{path}', 'path', 'sites/marketing/north')
    ).toBeUndefined();
  });

  it('encodes quoted breakout syntax and rejects unsafe integer or relative-path syntax', () => {
    expect(
      prepareUnencodedPathParameter("/drives/{drive-id}/search(q='{q}')", 'q', "report')/children")
    ).toBe('report%27%27)%2Fchildren');
    expect(
      prepareUnencodedPathParameter(
        "/workbook/worksheets/{id}/range(address='{address}')",
        'address',
        'A1)/tables?x=1'
      )
    ).toBe('A1)%2Ftables%3Fx%3D1');
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

  it('accepts only an ISO 8601 date-time in an unquoted DateTime argument', () => {
    const pattern =
      "/me/adhocCalls/getAllTranscripts(userId='{userId}',startDateTime={startDateTime},endDateTime={endDateTime})";
    expect(prepareUnencodedPathParameter(pattern, 'startDateTime', '2026-10-01T00:00:00Z')).toBe(
      '2026-10-01T00%3A00%3A00Z'
    );
    expect(prepareUnencodedPathParameter(pattern, 'endDateTime', '2026-10-08T12:30:00+02:00')).toBe(
      '2026-10-08T12%3A30%3A00%2B02%3A00'
    );
    for (const value of [
      '2026-10-01',
      '2026-10-01T00:00:00Z)/x',
      "2026-10-01T00:00:00Z,userId='x'",
    ]) {
      expect(unencodedPathParameterError(pattern, 'startDateTime', value)).toMatch(/ISO 8601/);
    }
  });
});
