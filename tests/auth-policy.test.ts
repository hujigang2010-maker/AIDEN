import { describe, expect, it } from 'vitest';

import {
  CHROME_USER_AGENT,
  decideWindowOpen,
  isGeminiNavigationUrl,
  isGoogleAuthUrl,
  isGoogleHost,
  sanitizeRequestHeaders,
  shouldReloadAfterAuthNavigation,
} from '../src/auth-policy';

describe('登录域名判断', () => {
  it('识别 Google 与 Gemini 主机', () => {
    expect(isGoogleHost('accounts.google.com')).toBe(true);
    expect(isGoogleHost('gemini.google.com')).toBe(true);
    expect(isGoogleHost('oauth2.googleapis.com')).toBe(true);
    expect(isGoogleHost('evil.example')).toBe(false);
  });

  it('识别授权页与 Gemini 回跳', () => {
    expect(isGoogleAuthUrl('https://accounts.google.com/v3/signin/identifier')).toBe(true);
    expect(isGeminiNavigationUrl('https://gemini.google.com/app')).toBe(true);
    expect(isGeminiNavigationUrl('https://example.com/app')).toBe(false);
  });
});

describe('弹窗策略', () => {
  it('Google 登录弹窗留在同一会话，外链拒绝', () => {
    expect(decideWindowOpen('https://accounts.google.com/signin')).toBe('allow-same-session');
    expect(decideWindowOpen('https://gemini.google.com/app')).toBe('allow-same-session');
    expect(decideWindowOpen('https://phishing.example/login')).toBe('deny');
    expect(decideWindowOpen('not-a-url')).toBe('deny');
  });

  it('授权完成后回到 Gemini 才刷新主窗口', () => {
    expect(shouldReloadAfterAuthNavigation('https://gemini.google.com/app')).toBe(true);
    expect(shouldReloadAfterAuthNavigation('https://accounts.google.com/signin')).toBe(false);
    expect(shouldReloadAfterAuthNavigation('https://gemini.google.com/signin/oops')).toBe(false);
  });
});

describe('请求头伪装', () => {
  it('覆盖 User-Agent 并去掉 Electron 指纹', () => {
    const headers = sanitizeRequestHeaders({
      Accept: 'text/html',
      'X-Electron': '1',
      'User-Agent': 'Electron/37.0',
    });

    expect(headers['User-Agent']).toBe(CHROME_USER_AGENT);
    expect(headers['X-Electron']).toBeUndefined();
    expect(headers.Accept).toBe('text/html');
    expect(headers['Sec-CH-UA']).toContain('Google Chrome');
    expect(CHROME_USER_AGENT.includes('Electron')).toBe(false);
  });
});
