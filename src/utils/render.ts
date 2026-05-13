import { marked, Renderer } from 'marked';
import hljs from 'highlight.js/lib/core';
import javascript from 'highlight.js/lib/languages/javascript';
import typescript from 'highlight.js/lib/languages/typescript';
import xml from 'highlight.js/lib/languages/xml';
import css from 'highlight.js/lib/languages/css';
import json from 'highlight.js/lib/languages/json';
import bash from 'highlight.js/lib/languages/bash';
import shell from 'highlight.js/lib/languages/shell';
import python from 'highlight.js/lib/languages/python';
import java from 'highlight.js/lib/languages/java';
import go from 'highlight.js/lib/languages/go';
import rust from 'highlight.js/lib/languages/rust';
import sql from 'highlight.js/lib/languages/sql';
import markdown from 'highlight.js/lib/languages/markdown';
import yaml from 'highlight.js/lib/languages/yaml';
import c from 'highlight.js/lib/languages/c';
import cpp from 'highlight.js/lib/languages/cpp';
import csharp from 'highlight.js/lib/languages/csharp';
import php from 'highlight.js/lib/languages/php';
import ruby from 'highlight.js/lib/languages/ruby';
import swift from 'highlight.js/lib/languages/swift';
import kotlin from 'highlight.js/lib/languages/kotlin';
import diff from 'highlight.js/lib/languages/diff';
import type { StyleMap, Theme } from '../themes';

hljs.registerLanguage('javascript', javascript);
hljs.registerLanguage('js', javascript);
hljs.registerLanguage('jsx', javascript);
hljs.registerLanguage('typescript', typescript);
hljs.registerLanguage('ts', typescript);
hljs.registerLanguage('tsx', typescript);
hljs.registerLanguage('xml', xml);
hljs.registerLanguage('html', xml);
hljs.registerLanguage('css', css);
hljs.registerLanguage('json', json);
hljs.registerLanguage('bash', bash);
hljs.registerLanguage('sh', bash);
hljs.registerLanguage('shell', shell);
hljs.registerLanguage('python', python);
hljs.registerLanguage('py', python);
hljs.registerLanguage('java', java);
hljs.registerLanguage('go', go);
hljs.registerLanguage('rust', rust);
hljs.registerLanguage('rs', rust);
hljs.registerLanguage('sql', sql);
hljs.registerLanguage('markdown', markdown);
hljs.registerLanguage('md', markdown);
hljs.registerLanguage('yaml', yaml);
hljs.registerLanguage('yml', yaml);
hljs.registerLanguage('c', c);
hljs.registerLanguage('cpp', cpp);
hljs.registerLanguage('c++', cpp);
hljs.registerLanguage('csharp', csharp);
hljs.registerLanguage('cs', csharp);
hljs.registerLanguage('php', php);
hljs.registerLanguage('ruby', ruby);
hljs.registerLanguage('rb', ruby);
hljs.registerLanguage('swift', swift);
hljs.registerLanguage('kotlin', kotlin);
hljs.registerLanguage('kt', kotlin);
hljs.registerLanguage('diff', diff);

function styleString(map: StyleMap): string {
  return Object.entries(map)
    .map(([k, v]) => `${k}:${v}`)
    .join(';');
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildRenderer(theme: Theme): Renderer {
  const s = theme.styles;
  const renderer = new Renderer();

  renderer.heading = ({ tokens, depth }) => {
    const text = renderer.parser.parseInline(tokens);
    const styleKey = `h${depth}` as keyof typeof s;
    const style = styleString(s[styleKey] as StyleMap);
    return `<h${depth} style="${style}">${text}</h${depth}>`;
  };

  renderer.paragraph = ({ tokens }) => {
    const text = renderer.parser.parseInline(tokens);
    return `<p style="${styleString(s.p)}">${text}</p>`;
  };

  renderer.strong = ({ tokens }) => {
    const text = renderer.parser.parseInline(tokens);
    return `<strong style="${styleString(s.strong)}">${text}</strong>`;
  };

  renderer.em = ({ tokens }) => {
    const text = renderer.parser.parseInline(tokens);
    return `<em style="${styleString(s.em)}">${text}</em>`;
  };

  renderer.del = ({ tokens }) => {
    const text = renderer.parser.parseInline(tokens);
    return `<del style="${styleString(s.del)}">${text}</del>`;
  };

  renderer.link = ({ href, title, tokens }) => {
    const text = renderer.parser.parseInline(tokens);
    const safeHref = (href || '').replace(/"/g, '&quot;');
    const t = title ? ` title="${escapeHtml(title)}"` : '';
    return `<a href="${safeHref}"${t} style="${styleString(s.a)}">${text}</a>`;
  };

  renderer.list = (token) => {
    const tag = token.ordered ? 'ol' : 'ul';
    const style = styleString(token.ordered ? s.ol : s.ul);
    const startAttr = token.ordered && token.start !== 1 ? ` start="${token.start}"` : '';
    let body = '';
    for (const item of token.items) {
      body += renderer.listitem(item);
    }
    return `<${tag} style="${style}"${startAttr}>${body}</${tag}>`;
  };

  renderer.listitem = (item) => {
    let body = '';
    if (item.task) {
      const checkbox = `<input type="checkbox" disabled${
        item.checked ? ' checked' : ''
      } style="margin-right:0.4em;" />`;
      body += checkbox;
    }
    body += renderer.parser.parse(item.tokens);
    return `<li style="${styleString(s.li)}">${body}</li>`;
  };

  renderer.blockquote = ({ tokens }) => {
    const inner = renderer.parser.parse(tokens);
    const withBlockquoteP = inner.replace(
      /<p style="[^"]*"/g,
      `<p style="${styleString(s.blockquoteP)}"`,
    );
    return `<blockquote style="${styleString(s.blockquote)}">${withBlockquoteP}</blockquote>`;
  };

  renderer.hr = () => `<hr style="${styleString(s.hr)}" />`;

  renderer.codespan = ({ text }) => {
    return `<code style="${styleString(s.code)}">${text}</code>`;
  };

  renderer.code = ({ text, lang }) => {
    let highlighted = escapeHtml(text);
    const language = (lang || '').match(/^\S*/)?.[0] || '';
    if (language && hljs.getLanguage(language)) {
      try {
        highlighted = hljs.highlight(text, { language }).value;
      } catch {
        highlighted = escapeHtml(text);
      }
    } else if (text) {
      try {
        highlighted = hljs.highlightAuto(text).value;
      } catch {
        highlighted = escapeHtml(text);
      }
    }
    return `<pre style="${styleString(s.pre)}"><code style="${styleString(
      s.preCode,
    )}" class="hljs language-${language}">${highlighted}</code></pre>`;
  };

  renderer.image = ({ href, title, text }) => {
    const safeHref = (href || '').replace(/"/g, '&quot;');
    const t = title ? ` title="${escapeHtml(title)}"` : '';
    const alt = text ? ` alt="${escapeHtml(text)}"` : '';
    if (text) {
      return `<figure style="${styleString(s.figure)}"><img src="${safeHref}"${alt}${t} style="${styleString(
        s.img,
      )}" /><figcaption style="${styleString(s.figcaption)}">${escapeHtml(text)}</figcaption></figure>`;
    }
    return `<img src="${safeHref}"${alt}${t} style="${styleString(s.img)}" />`;
  };

  renderer.table = (token) => {
    let headerRow = '';
    for (let i = 0; i < token.header.length; i++) {
      const cell = token.header[i];
      const align = cell.align ? `text-align:${cell.align};` : '';
      const cellHtml = renderer.parser.parseInline(cell.tokens);
      headerRow += `<th style="${styleString(s.th)}${align}">${cellHtml}</th>`;
    }
    let bodyRows = '';
    for (const row of token.rows) {
      let rowHtml = '';
      for (const cell of row) {
        const align = cell.align ? `text-align:${cell.align};` : '';
        const cellHtml = renderer.parser.parseInline(cell.tokens);
        rowHtml += `<td style="${styleString(s.td)}${align}">${cellHtml}</td>`;
      }
      bodyRows += `<tr style="${styleString(s.tr)}">${rowHtml}</tr>`;
    }
    return `<table style="${styleString(s.table)}"><thead style="${styleString(
      s.thead,
    )}"><tr>${headerRow}</tr></thead><tbody>${bodyRows}</tbody></table>`;
  };

  return renderer;
}

export interface RenderResult {
  containerStyle: string;
  bodyHtml: string;
  fullHtml: string;
}

export function renderMarkdown(markdown: string, theme: Theme): RenderResult {
  const renderer = buildRenderer(theme);
  marked.setOptions({
    gfm: true,
    breaks: false,
  });
  const html = marked.parse(markdown, { renderer, async: false }) as string;
  const containerStyle = Object.entries(theme.styles.container)
    .map(([k, v]) => `${k}:${v}`)
    .join(';');
  const fullHtml = `<section style="${containerStyle}">${html}</section>`;
  return { containerStyle, bodyHtml: html, fullHtml };
}
