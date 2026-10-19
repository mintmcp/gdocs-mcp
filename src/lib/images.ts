/**
 * Image URLs for insertInlineImage / replaceImage. Google fetches the image
 * itself, once, so the URL has to be reachable without a login and short
 * enough for the API: the checks here fail fast with a message that says why.
 */

export const MAX_IMAGE_URL_LENGTH = 2048;

export const IMAGE_LIMITS_NOTE =
  'Google fetches the image itself: the URL must be https and reachable without a login, at most 2 kB long, ' +
  'and the image PNG, JPEG or GIF, under 50 MB and 25 megapixels.';

const DRIVE_FILE_LINK = /^https:\/\/drive\.google\.com\/(?:file\/d\/([\w-]+)|open\?(?:.*&)?id=([\w-]+)|uc\?(?:.*&)?id=([\w-]+))/;

/**
 * The URL Google should fetch. A Drive share link points at a viewer page, so
 * it becomes the file's download URL; that only works when the file is shared
 * with anyone who has the link.
 */
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
  const drive = DRIVE_FILE_LINK.exec(url);
  const resolved = drive
    ? `https://drive.google.com/uc?export=download&id=${drive[1] ?? drive[2] ?? drive[3]}`
    : url;
  if (resolved.length > MAX_IMAGE_URL_LENGTH) {
    throw new Error(`image_url is ${resolved.length} characters; Google accepts at most ${MAX_IMAGE_URL_LENGTH} (2 kB). Host the image at a shorter URL.`);
  }
  return resolved;
}

/** objectSize for insertInlineImage; a single dimension keeps the aspect ratio. */
export function imageObjectSize(widthPt?: number, heightPt?: number): { objectSize: Record<string, unknown> } | {} {
  const objectSize: Record<string, unknown> = {};
  if (widthPt !== undefined) objectSize.width = { magnitude: widthPt, unit: 'PT' };
  if (heightPt !== undefined) objectSize.height = { magnitude: heightPt, unit: 'PT' };
  return Object.keys(objectSize).length ? { objectSize } : {};
}
