import { useEffect, useMemo, useRef, useState } from 'react';
import './App.css';
import { Toolbar } from './components/Toolbar';
import { Editor } from './components/Editor';
import { Preview } from './components/Preview';
import { themes, getTheme } from './themes';
import { renderMarkdown } from './utils/render';
import { copyHtmlToClipboard } from './utils/clipboard';
import { inlineHljsStyles } from './utils/inlineHljs';
import { SAMPLE_MD } from './sample';

const STORAGE_KEY = 'mp-typer:content';
const THEME_KEY = 'mp-typer:theme';

function App() {
  const [markdown, setMarkdown] = useState<string>(() => {
    if (typeof window === 'undefined') return SAMPLE_MD;
    return window.localStorage.getItem(STORAGE_KEY) ?? SAMPLE_MD;
  });
  const [themeId, setThemeId] = useState<string>(() => {
    if (typeof window === 'undefined') return 'default';
    return window.localStorage.getItem(THEME_KEY) ?? 'default';
  });
  const [toast, setToast] = useState<string>('');

  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const previewRef = useRef<HTMLDivElement>(null);

  const theme = useMemo(() => getTheme(themeId), [themeId]);
  const rendered = useMemo(() => renderMarkdown(markdown, theme), [markdown, theme]);

  useEffect(() => {
    try {
      window.localStorage.setItem(STORAGE_KEY, markdown);
    } catch {
      // ignore quota errors
    }
  }, [markdown]);

  useEffect(() => {
    try {
      window.localStorage.setItem(THEME_KEY, themeId);
    } catch {
      // ignore
    }
  }, [themeId]);

  useEffect(() => {
    if (previewRef.current) {
      inlineHljsStyles(previewRef.current);
    }
  }, [rendered]);

  function showToast(msg: string) {
    setToast(msg);
    window.setTimeout(() => setToast(''), 2200);
  }

  async function handleCopy() {
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = rendered.fullHtml;
    inlineHljsStyles(tempDiv);
    const html = tempDiv.innerHTML;
    const ok = await copyHtmlToClipboard(html, markdown);
    showToast(ok ? '✅ 已复制富文本，粘贴到公众号编辑器即可' : '❌ 复制失败，请手动复制预览区');
  }

  async function handleCopyHtml() {
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = rendered.fullHtml;
    inlineHljsStyles(tempDiv);
    const html = tempDiv.innerHTML;
    try {
      await navigator.clipboard.writeText(html);
      showToast('✅ HTML 源码已复制到剪贴板');
    } catch {
      showToast('❌ 复制失败');
    }
  }

  function handleClear() {
    if (window.confirm('确定要清空当前内容吗？')) {
      setMarkdown('');
      requestAnimationFrame(() => textareaRef.current?.focus());
    }
  }

  function handleLoadSample() {
    setMarkdown(SAMPLE_MD);
  }

  function handleImportFile(file: File) {
    const reader = new FileReader();
    reader.onload = () => {
      setMarkdown(String(reader.result ?? ''));
    };
    reader.readAsText(file, 'utf-8');
  }

  function handleExport() {
    const blob = new Blob([markdown], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `article-${new Date().toISOString().slice(0, 10)}.md`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-logo">公</span>
          <div>
            <div className="brand-title">MP Typer · 公众号排版助手</div>
            <div className="brand-sub">Markdown · 实时预览 · 一键复制</div>
          </div>
        </div>

        <div className="theme-picker">
          <label className="topbar-label">主题</label>
          <select
            className="theme-select"
            value={themeId}
            onChange={(e) => setThemeId(e.target.value)}
            title={theme.description}
          >
            {themes.map((t) => (
              <option key={t.id} value={t.id}>
                {t.name}
              </option>
            ))}
          </select>
        </div>

        <div className="topbar-actions">
          <label className="btn btn-ghost">
            导入
            <input
              type="file"
              accept=".md,.markdown,.txt,text/markdown,text/plain"
              hidden
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) handleImportFile(f);
                e.target.value = '';
              }}
            />
          </label>
          <button className="btn btn-ghost" onClick={handleExport}>
            导出 .md
          </button>
          <button className="btn btn-ghost" onClick={handleLoadSample}>
            示例
          </button>
          <button className="btn btn-ghost" onClick={handleClear}>
            清空
          </button>
          <button className="btn btn-ghost" onClick={handleCopyHtml} title="复制带内联样式的 HTML 源码">
            复制 HTML
          </button>
          <button className="btn btn-primary" onClick={handleCopy}>
            复制到公众号
          </button>
        </div>
      </header>

      <main className="workspace">
        <section className="pane editor-pane">
          <div className="pane-head">
            <span>编辑 · Markdown</span>
            <span className="pane-hint">支持 Ctrl/Cmd+B 加粗 · Ctrl/Cmd+I 斜体</span>
          </div>
          <Toolbar textareaRef={textareaRef} value={markdown} onChange={setMarkdown} />
          <Editor value={markdown} onChange={setMarkdown} textareaRef={textareaRef} />
        </section>

        <section className="pane preview-pane">
          <div className="pane-head">
            <span>预览 · {theme.name}</span>
            <span className="pane-hint">{theme.description}</span>
          </div>
          <Preview ref={previewRef} fullHtml={rendered.fullHtml} />
        </section>
      </main>

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}

export default App;
