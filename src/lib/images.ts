import { pt } from './docRequests.js';

export const MAX_IMAGE_URL_LENGTH = 2048;

export const IMAGE_FETCH_REFUSAL = /image|retriev|fetch|url|uri/i;

export const IMAGE_LIMITS_NOTE =
  'Google fetches the image itself: the URL must be https and reachable without a login, at most 2 kB long, ' +
  'and the image PNG, JPEG or GIF, under 50 MB and 25 megapixels.';

function driveFileId(url: URL): string | undefined {
  if (url.hostname !== 'drive.google.com' && url.hostname !== 'docs.google.com') return undefined;
  const id = /^\/file(?:\/u\/\d+)?\/d\/([\w-]+)/.exec(url.pathname)?.[1]
    ?? (url.pathname === '/open' || url.pathname === '/uc' ? url.searchParams.get('id') : null);
  return id && /^[\w-]+$/.test(id) ? id : undefined;
}

export function imageSourceUrl(url: string): string {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    throw new Error(`image_url must be an absolute https URL, received '${url.slice(0, 60)}'.`);
  }
  if (parsed.protocol !== 'https:') {
    throw new Error(`image_url must use https (Google fetches it from the public web; data: URLs and plain http are refused), received '${url.slice(0, 60)}'.`);
  }
  const driveId = driveFileId(parsed);
  const resourceKey = parsed.searchParams.get('resourcekey');
  const resolved = driveId
    ? `https://drive.google.com/uc?export=download&id=${driveId}${resourceKey ? `&resourcekey=${encodeURIComponent(resourceKey)}` : ''}`
    : url;
  if (resolved.length > MAX_IMAGE_URL_LENGTH) {
    throw new Error(`image_url is ${resolved.length} characters; Google accepts at most ${MAX_IMAGE_URL_LENGTH} (2 kB). Host the image at a shorter URL.`);
  }
  return resolved;
}

export function imageObjectSize(widthPt?: number, heightPt?: number): { objectSize: Record<string, unknown> } | {} {
  const objectSize: Record<string, unknown> = {};
  if (widthPt !== undefined) objectSize.width = pt(widthPt);
  if (heightPt !== undefined) objectSize.height = pt(heightPt);
  return Object.keys(objectSize).length ? { objectSize } : {};
}
