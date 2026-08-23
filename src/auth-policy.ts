/**
 * Gemini 桌面端登录策略。
 *
 * 旧版包装器把 https://gemini.google.com/app 塞进 iframe，并剥离
 * X-Frame-Options。浏览器里 Google 授权可以成功，但应用内会话
 * 写不回顶层窗口，表现为「网络和授权都正常，却始终登不进去」。
 *
 * 这里强制：顶层导航、同一持久化分区、只允许 Google 登录弹窗。
 */

export const GEMINI_APP_URL = 'https://gemini.google.com/app';
export const SESSION_PARTITION = 'persist:gemini';

export const CHROME_USER_AGENT =
  'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.7778.96 Safari/537.36';

const GOOGLE_HOST_SUFFIXES = [
  'google.com',
  'googleusercontent.com',
  'gstatic.com',
  'googleapis.com',
  'ggpht.com',
] as const;

export type WindowOpenDecision = 'allow-same-session' | 'deny';

export function hostnameOf(url: string): string | null {
  try {
    return new URL(url).hostname.toLowerCase();
  } catch {
    return null;
  }
}

export function isGoogleHost(hostname: string): boolean {
  const host = hostname.toLowerCase();
  return GOOGLE_HOST_SUFFIXES.some((suffix) => host === suffix || host.endsWith(`.${suffix}`));
}

export function isGeminiNavigationUrl(url: string): boolean {
  const hostname = hostnameOf(url);
  if (!hostname) {
    return false;
  }
  return hostname === 'gemini.google.com' || hostname.endsWith('.gemini.google.com');
}

export function isGoogleAuthUrl(url: string): boolean {
  const hostname = hostnameOf(url);
  if (!hostname) {
    return false;
  }
  return (
    hostname === 'accounts.google.com' ||
    hostname.endsWith('.accounts.google.com') ||
    hostname === 'accounts.youtube.com' ||
    hostname === 'ogs.google.com'
  );
}

/**
 * 登录相关弹窗必须留在同一 persist:gemini 分区。
 * 外链一律拒绝，避免授权跳到系统浏览器后 cookie 回不来。
 */
export function decideWindowOpen(url: string): WindowOpenDecision {
  const hostname = hostnameOf(url);
  if (!hostname) {
    return 'deny';
  }
  if (isGoogleHost(hostname)) {
    return 'allow-same-session';
  }
  return 'deny';
}

export function shouldReloadAfterAuthNavigation(url: string): boolean {
  return isGeminiNavigationUrl(url) && !url.includes('/signin/');
}

export function buildChromeClientHints(): Record<string, string> {
  return {
    'User-Agent': CHROME_USER_AGENT,
    'Sec-CH-UA': '"Google Chrome";v="148", "Chromium";v="148", "Not.A/Brand";v="99"',
    'Sec-CH-UA-Mobile': '?0',
    'Sec-CH-UA-Platform': '"Linux"',
  };
}

export function sanitizeRequestHeaders(
  headers: Record<string, string>,
): Record<string, string> {
  const next: Record<string, string> = {};
  for (const [key, value] of Object.entries(headers)) {
    if (key.toLowerCase() === 'x-electron') {
      continue;
    }
    next[key] = value;
  }
  Object.assign(next, buildChromeClientHints());
  return next;
}
