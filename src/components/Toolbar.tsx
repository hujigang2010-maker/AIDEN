import type { RefObject } from 'react';

interface ToolbarProps {
  textareaRef: RefObject<HTMLTextAreaElement | null>;
  value: string;
  onChange: (value: string) => void;
}

type Action =
  | { type: 'wrap'; label: string; title: string; before: string; after: string; placeholder?: string }
  | { type: 'linePrefix'; label: string; title: string; prefix: string; placeholder?: string }
  | { type: 'insert'; label: string; title: string; text: string };

const ACTIONS: Action[] = [
  { type: 'linePrefix', label: 'H1', title: '一级标题', prefix: '# ', placeholder: '一级标题' },
  { type: 'linePrefix', label: 'H2', title: '二级标题', prefix: '## ', placeholder: '二级标题' },
  { type: 'linePrefix', label: 'H3', title: '三级标题', prefix: '### ', placeholder: '三级标题' },
  { type: 'wrap', label: 'B', title: '加粗 (Ctrl/Cmd+B)', before: '**', after: '**', placeholder: '加粗文字' },
  { type: 'wrap', label: 'I', title: '斜体 (Ctrl/Cmd+I)', before: '*', after: '*', placeholder: '斜体文字' },
  { type: 'wrap', label: 'S', title: '删除线', before: '~~', after: '~~', placeholder: '删除线' },
  { type: 'wrap', label: '`', title: '行内代码', before: '`', after: '`', placeholder: 'code' },
  { type: 'linePrefix', label: '“”', title: '引用', prefix: '> ', placeholder: '引用内容' },
  { type: 'linePrefix', label: '•', title: '无序列表', prefix: '- ', placeholder: '列表项' },
  { type: 'linePrefix', label: '1.', title: '有序列表', prefix: '1. ', placeholder: '列表项' },
  {
    type: 'insert',
    label: '链接',
    title: '插入链接',
    text: '[链接文字](https://)',
  },
  {
    type: 'insert',
    label: '图片',
    title: '插入图片链接',
    text: '![图片描述](https://)',
  },
  {
    type: 'insert',
    label: '代码块',
    title: '插入代码块',
    text: '\n```js\n// code here\n```\n',
  },
  {
    type: 'insert',
    label: '表格',
    title: '插入表格',
    text:
      '\n| 列 1 | 列 2 | 列 3 |\n| --- | --- | --- |\n| 内容 | 内容 | 内容 |\n| 内容 | 内容 | 内容 |\n',
  },
  { type: 'insert', label: '---', title: '分割线', text: '\n\n---\n\n' },
];

export function Toolbar({ textareaRef, value, onChange }: ToolbarProps) {
  function applyWrap(before: string, after: string, placeholder = '') {
    const ta = textareaRef.current;
    if (!ta) return;
    const start = ta.selectionStart;
    const end = ta.selectionEnd;
    const selected = value.slice(start, end) || placeholder;
    const next = value.slice(0, start) + before + selected + after + value.slice(end);
    onChange(next);
    const cursorStart = start + before.length;
    const cursorEnd = cursorStart + selected.length;
    requestAnimationFrame(() => {
      ta.focus();
      ta.setSelectionRange(cursorStart, cursorEnd);
    });
  }

  function applyLinePrefix(prefix: string, placeholder = '') {
    const ta = textareaRef.current;
    if (!ta) return;
    const start = ta.selectionStart;
    const end = ta.selectionEnd;
    const before = value.slice(0, start);
    const lineStart = before.lastIndexOf('\n') + 1;
    const after = value.slice(end);
    const segment = value.slice(lineStart, end);
    const lines = segment.length === 0 ? [placeholder] : segment.split('\n');
    const transformed = lines.map((l) => prefix + l).join('\n');
    const next = value.slice(0, lineStart) + transformed + after;
    onChange(next);
    const newPos = lineStart + transformed.length;
    requestAnimationFrame(() => {
      ta.focus();
      ta.setSelectionRange(newPos, newPos);
    });
  }

  function applyInsert(text: string) {
    const ta = textareaRef.current;
    if (!ta) return;
    const start = ta.selectionStart;
    const end = ta.selectionEnd;
    const next = value.slice(0, start) + text + value.slice(end);
    onChange(next);
    const newPos = start + text.length;
    requestAnimationFrame(() => {
      ta.focus();
      ta.setSelectionRange(newPos, newPos);
    });
  }

  function handleAction(action: Action) {
    if (action.type === 'wrap') applyWrap(action.before, action.after, action.placeholder);
    else if (action.type === 'linePrefix') applyLinePrefix(action.prefix, action.placeholder);
    else applyInsert(action.text);
  }

  return (
    <div className="toolbar">
      {ACTIONS.map((a) => (
        <button
          key={a.label + a.title}
          type="button"
          className="toolbar-btn"
          title={a.title}
          onClick={() => handleAction(a)}
        >
          {a.label}
        </button>
      ))}
    </div>
  );
}
