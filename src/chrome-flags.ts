/** Linux 云桌面 / 无 GPU 环境需要与系统 Chrome 相同的兼容开关。 */
export const LINUX_CHROME_COMPAT_SWITCHES: Array<[string, string?]> = [
  ['no-sandbox'],
  ['disable-gpu-sandbox'],
  ['disable-dev-shm-usage'],
  ['use-gl', 'angle'],
  ['use-angle', 'swiftshader-webgl'],
  ['password-store', 'basic'],
];
