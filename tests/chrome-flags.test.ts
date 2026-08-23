import { describe, expect, it } from 'vitest';

import { LINUX_CHROME_COMPAT_SWITCHES } from '../src/chrome-flags';

describe('Linux Chrome 兼容开关', () => {
  it('包含无沙箱、禁用 /dev/shm 和 SwiftShader', () => {
    const names = LINUX_CHROME_COMPAT_SWITCHES.map(([name]) => name);
    expect(names).toContain('no-sandbox');
    expect(names).toContain('disable-dev-shm-usage');
    expect(names).toContain('use-angle');
    const angle = LINUX_CHROME_COMPAT_SWITCHES.find(([name]) => name === 'use-angle');
    expect(angle?.[1]).toBe('swiftshader-webgl');
  });
});
