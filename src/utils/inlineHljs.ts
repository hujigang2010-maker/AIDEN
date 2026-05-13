const HLJS_COLOR_MAP: Record<string, string> = {
  'hljs-comment': '#8b949e;font-style:italic',
  'hljs-quote': '#8b949e;font-style:italic',
  'hljs-keyword': '#ff7b72',
  'hljs-selector-tag': '#ff7b72',
  'hljs-literal': '#79c0ff',
  'hljs-section': '#ff7b72',
  'hljs-link': '#79c0ff',
  'hljs-string': '#a5d6ff',
  'hljs-title': '#d2a8ff',
  'hljs-name': '#7ee787',
  'hljs-type': '#ffa657',
  'hljs-attribute': '#79c0ff',
  'hljs-symbol': '#79c0ff',
  'hljs-bullet': '#79c0ff',
  'hljs-built_in': '#ffa657',
  'hljs-addition': '#aff5b4;background:#033a16',
  'hljs-variable': '#ffa657',
  'hljs-template-tag': '#ff7b72',
  'hljs-template-variable': '#ffa657',
  'hljs-comment2': '#8b949e',
  'hljs-deletion': '#ffdcd7;background:#67060c',
  'hljs-meta': '#8b949e',
  'hljs-number': '#79c0ff',
  'hljs-params': '#ffa657',
  'hljs-tag': '#7ee787',
  'hljs-class': '#d2a8ff',
  'hljs-function': '#d2a8ff',
  'hljs-property': '#79c0ff',
  'hljs-regexp': '#a5d6ff',
  'hljs-operator': '#ff7b72',
  'hljs-punctuation': '#c9d1d9',
  'hljs-doctag': '#7ee787',
  'hljs-emphasis': 'font-style:italic',
  'hljs-strong': 'font-weight:bold',
  'hljs-formula': '#79c0ff',
  'hljs-subst': '#c9d1d9',
};

export function inlineHljsStyles(rootEl: HTMLElement) {
  const spans = rootEl.querySelectorAll<HTMLElement>('span[class*="hljs-"]');
  spans.forEach((span) => {
    const classes = Array.from(span.classList).filter((c) => c.startsWith('hljs-'));
    if (classes.length === 0) return;
    const styles: string[] = [];
    for (const cls of classes) {
      const v = HLJS_COLOR_MAP[cls];
      if (v) {
        if (v.includes(':')) {
          styles.push(v);
        } else {
          styles.push(`color:${v}`);
        }
      }
    }
    if (styles.length > 0) {
      const existing = span.getAttribute('style') || '';
      span.setAttribute('style', existing + (existing ? ';' : '') + styles.join(';'));
    }
  });
}
