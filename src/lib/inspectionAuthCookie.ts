export type InspectionAuthRole = "admin" | "operator" | "viewer";

export interface InspectionAuthUserCookie {
  id: string;
  username: string;
  displayName: string;
  role: InspectionAuthRole;
  lastLoginAt: string | null;
}

export interface InspectionAuthCookie {
  user: InspectionAuthUserCookie;
  issuedAt: number;
}

export const INSPECTION_AUTH_COOKIE_NAME = "inspection_auth_hint";
export const INSPECTION_AUTH_REVALIDATION_MS = 10 * 60 * 1000;

function browserCookieString() {
  try {
    const documentLike = (globalThis as typeof globalThis & { document?: { cookie?: string } }).document;
    return documentLike?.cookie || "";
  } catch {
    return "";
  }
}

function isRole(value: unknown): value is InspectionAuthRole {
  return value === "admin" || value === "operator" || value === "viewer";
}

function isUserCookie(value: unknown): value is InspectionAuthUserCookie {
  if (!value || typeof value !== "object") return false;
  const user = value as Record<string, unknown>;
  return typeof user.id === "string"
    && typeof user.username === "string"
    && typeof user.displayName === "string"
    && isRole(user.role)
    && (typeof user.lastLoginAt === "string" || user.lastLoginAt === null);
}

function decodeBase64Url(value: string) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(value.length / 4) * 4, "=");
  const binary = atob(normalized);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function isAuthCookie(value: unknown): value is InspectionAuthCookie {
  if (!value || typeof value !== "object") return false;
  const cookie = value as Record<string, unknown>;
  return isUserCookie(cookie.user) && Number.isFinite(cookie.issuedAt);
}

export function readInspectionAuthCookie(cookieString = browserCookieString()): InspectionAuthCookie | null {
  const raw = cookieString
    .split(";")
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${INSPECTION_AUTH_COOKIE_NAME}=`))
    ?.slice(`${INSPECTION_AUTH_COOKIE_NAME}=`.length);
  if (!raw) return null;
  try {
    const parsed: unknown = JSON.parse(decodeBase64Url(decodeURIComponent(raw)));
    return isAuthCookie(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

export function clearInspectionAuthCookie() {
  try {
    const documentLike = (globalThis as typeof globalThis & { document?: { cookie?: string } }).document;
    if (documentLike) documentLike.cookie = `${INSPECTION_AUTH_COOKIE_NAME}=; Max-Age=0; Path=/; SameSite=Lax`;
  } catch {
    // Cookie cleanup is best effort; the server clears it on the next auth response.
  }
}

export function isInspectionAuthCookieStale(cookie: InspectionAuthCookie, now = Date.now()) {
  return !Number.isFinite(cookie.issuedAt) || now - cookie.issuedAt >= INSPECTION_AUTH_REVALIDATION_MS;
}
