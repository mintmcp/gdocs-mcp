// Low-level Google Drive / Docs HTTP helpers, shared by tools.ts and the
// vendored label module

export const GOOGLE_DRIVE_API = 'https://www.googleapis.com/drive/v3';
export const GOOGLE_DOCS_API = 'https://docs.googleapis.com/v1/documents';

/**
 * Error thrown by Google API helpers. Carries enough structured detail
 * (status, Google's `error.status` enum, retry-after) for callers to surface
 * machine-readable error envelopes instead of opaque strings.
 */
export class GoogleApiError extends Error {
  status: number;
  code?: string;
  retryAfter?: number;
  api: 'drive' | 'docs';
  // Google's `error.details[]` payload (e.g. `BadRequest.fieldViolations`,
  // `Help`, request-index hints). Surfaced so callers can pinpoint which
  // request in a multi-request batchUpdate actually failed.
  details?: unknown[];

  constructor(message: string, status: number, api: 'drive' | 'docs', opts: { code?: string; retryAfter?: number; details?: unknown[] } = {}) {
    super(message);
    this.name = 'GoogleApiError';
    this.status = status;
    this.api = api;
    this.code = opts.code;
    this.retryAfter = opts.retryAfter;
    this.details = opts.details;
  }
}

export function safeJsonParse(text: string): any | null {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

/**
 * Parse a non-OK Google API response into a `GoogleApiError`.
 *
 * Note: Google's Drive/Docs APIs return 404 for both "doesn't exist" and
 * "you don't have access" — we surface a disambiguated message so callers
 * don't make wrong assumptions.
 */
export async function buildGoogleApiError(
  response: Response,
  api: 'drive' | 'docs'
): Promise<GoogleApiError> {
  const errorText = await response.text().catch(() => '');
  const errorJson = errorText ? safeJsonParse(errorText) : null;
  const googleMessage: string | undefined = errorJson?.error?.message;
  const googleCode: string | undefined = errorJson?.error?.status;
  const googleDetails: unknown[] | undefined = Array.isArray(errorJson?.error?.details) && errorJson.error.details.length > 0
    ? errorJson.error.details
    : undefined;

  let message: string;
  switch (response.status) {
    case 401:
      message = 'Authentication failed. Please re-authenticate.';
      break;
    case 403:
      message = googleMessage
        ? `Permission denied: ${googleMessage}`
        : `Permission denied. Make sure you have granted ${api === 'docs' ? 'Docs' : 'Drive'} access.`;
      break;
    case 404:
      message = api === 'docs'
        ? 'Document not found or you do not have permission to access it'
        : 'File or document not found or you do not have permission to access it';
      break;
    case 429:
      message = googleMessage || 'Rate limit exceeded. Retry after a short delay.';
      break;
    default:
      message = googleMessage || errorText || `Google ${api === 'docs' ? 'Docs' : 'Drive'} API error (${response.status})`;
  }

  let retryAfter: number | undefined;
  if (response.status === 429 || response.status === 503) {
    const header = response.headers.get('retry-after');
    if (header) {
      const parsed = parseInt(header, 10);
      if (!Number.isNaN(parsed) && parsed >= 0) retryAfter = parsed;
    }
  }

  return new GoogleApiError(message, response.status, api, { code: googleCode, retryAfter, details: googleDetails });
}


/**
 * Helper to make authenticated requests to Google Drive API
 */
export async function makeDriveRequest(
  endpoint: string,
  accessToken: string,
  options: RequestInit = {}
): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GOOGLE_DRIVE_API}${endpoint}`;

  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json',
      ...options.headers,
    },
  });

  if (!response.ok) {
    throw await buildGoogleApiError(response, 'drive');
  }

  return response.json();
}


/**
 * Helper to make authenticated requests to Google Docs API
 */
export async function makeDocsRequest(
  endpoint: string,
  accessToken: string,
  options: RequestInit = {}
): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GOOGLE_DOCS_API}${endpoint}`;

  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      ...options.headers,
    },
  });

  if (!response.ok) {
    throw await buildGoogleApiError(response, 'docs');
  }

  return response.json();
}
