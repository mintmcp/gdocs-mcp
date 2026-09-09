import { describe, test, expect } from 'vitest';
import { SCOPES, grantedScopes, expandScopes } from './scopes.js';
import { createServer } from './server.js';
import { GoogleDocsTools } from './tools.js';

describe('grantedScopes', () => {
  test('unset profile means unrestricted', () => {
    expect(grantedScopes(undefined)).toBeNull();
  });

  test('resolves a profile to its scope set', () => {
    expect(grantedScopes('standard')).toEqual(new Set([
      SCOPES.DRIVE_READONLY, SCOPES.DRIVE_FILE, SCOPES.DOCUMENTS,
    ]));
  });

  test('labels adds the label read scope on top of standard', () => {
    expect(grantedScopes('labels')).toEqual(new Set([
      SCOPES.DRIVE_READONLY, SCOPES.DRIVE_FILE, SCOPES.DOCUMENTS,
      SCOPES.DRIVE_LABELS_READONLY,
    ]));
  });

  test('an unknown profile fails at boot, not silently serving every tool', () => {
    expect(() => grantedScopes('readonly')).toThrow(/unknown profile/i);
  });

  test('an empty profile fails at boot too', () => {
    expect(() => grantedScopes('  ')).toThrow(/set but empty/i);
  });
});

function toolNames(granted: Set<string> | null): string[] {
  const server = createServer(granted);
  return Object.keys((server as any)._registeredTools ?? {}).sort();
}

const ALL_TOOLS = Object.keys(GoogleDocsTools.getTools()).sort();

describe('createServer tool surface', () => {
  test('unrestricted registers every tool', () => {
    expect(toolNames(null)).toEqual(ALL_TOOLS);
  });

  test('both profiles register every tool: the label scope gates enrichment, not a tool', () => {
    expect(toolNames(grantedScopes('standard'))).toEqual(ALL_TOOLS);
    expect(toolNames(grantedScopes('labels'))).toEqual(ALL_TOOLS);
  });

  test('a grant covering nothing registers nothing', () => {
    expect(toolNames(expandScopes([]))).toEqual([]);
  });

  test('labels.readonly alone registers nothing: no tool declares it', () => {
    expect(toolNames(expandScopes([SCOPES.DRIVE_LABELS_READONLY]))).toEqual([]);
  });
});
