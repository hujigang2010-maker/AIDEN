export async function copyHtmlToClipboard(html: string, plainText = ''): Promise<boolean> {
  if (
    typeof window !== 'undefined' &&
    window.isSecureContext &&
    navigator.clipboard &&
    typeof window.ClipboardItem !== 'undefined'
  ) {
    try {
      const htmlBlob = new Blob([html], { type: 'text/html' });
      const textBlob = new Blob([plainText || html.replace(/<[^>]+>/g, '')], {
        type: 'text/plain',
      });
      await navigator.clipboard.write([
        new ClipboardItem({ 'text/html': htmlBlob, 'text/plain': textBlob }),
      ]);
      return true;
    } catch {
      // fall back below
    }
  }

  return fallbackCopy(html);
}

function fallbackCopy(html: string): boolean {
  const container = document.createElement('div');
  container.contentEditable = 'true';
  container.style.position = 'fixed';
  container.style.top = '-10000px';
  container.style.left = '-10000px';
  container.style.opacity = '0';
  container.innerHTML = html;
  document.body.appendChild(container);

  const selection = window.getSelection();
  if (!selection) {
    document.body.removeChild(container);
    return false;
  }
  const range = document.createRange();
  range.selectNodeContents(container);
  selection.removeAllRanges();
  selection.addRange(range);

  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch {
    // ignore
  }
  selection.removeAllRanges();
  document.body.removeChild(container);
  return ok;
}
