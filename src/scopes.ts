export const SCOPES = {
  DRIVE_READONLY: "https://www.googleapis.com/auth/drive.readonly",
  DRIVE_FILE: "https://www.googleapis.com/auth/drive.file",
  DOCUMENTS: "https://www.googleapis.com/auth/documents",
  DRIVE_LABELS_READONLY: "https://www.googleapis.com/auth/drive.labels.readonly",
} as const;

// Editing a profile forces every user of its connector to re-consent, so
// frozen profiles never change. Keep each profile matching the scopes its
// brokered mcp-registry entry requests
export const PROFILES: Record<string, readonly string[]> = {
  "standard": [SCOPES.DRIVE_READONLY, SCOPES.DRIVE_FILE, SCOPES.DOCUMENTS],
  "labels": [
    SCOPES.DRIVE_READONLY,
    SCOPES.DRIVE_FILE,
    SCOPES.DOCUMENTS,
    SCOPES.DRIVE_LABELS_READONLY,
  ],
};

// Google's scope hierarchy, not our policy. Never add an implication Google
// doesn't grant; empty until a profile carries a superset scope
const IMPLIES: Record<string, readonly string[]> = {};

export function expandScopes(granted: Iterable<string>): Set<string> {
  const out = new Set<string>();
  for (const scope of granted) {
    out.add(scope);
    for (const implied of IMPLIES[scope] ?? []) out.add(implied);
  }
  return out;
}

/**
 * Unset means every tool registers (self-hosted default). An unknown or
 * empty profile throws so a misdeploy fails at boot, not silently
 */
export function grantedScopes(profile = process.env.PROFILE): Set<string> | null {
  if (profile === undefined) return null;

  const name = profile.trim();
  if (name === "") {
    throw new Error("PROFILE is set but empty. Unset it to register every tool.");
  }

  const scopes = PROFILES[name];
  if (!scopes) {
    throw new Error(
      `Unknown PROFILE "${name}". Known profiles: ${Object.keys(PROFILES).join(", ")}.`,
    );
  }
  return expandScopes(scopes);
}

export function isToolGranted(toolScope: string | undefined, granted: Set<string> | null): boolean {
  if (granted === null) return true;
  if (!toolScope) return true;
  return granted.has(toolScope);
}
