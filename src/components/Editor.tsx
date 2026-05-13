import { type RefObject, type KeyboardEvent } from 'react';

interface EditorProps {
  value: string;
  onChange: (v: string) => void;
  textareaRef: RefObject<HTMLTextAreaElement | null>;
}

export function Editor({ value, onChange, textareaRef }: EditorProps) {
  function onKeyDown(e: KeyboardEvent<HTMLTextAreaElement>) {
    const ta = e.currentTarget;

    if ((e.metaKey || e.ctrlKey) && (e.key === 'b' || e.key === 'B')) {
      e.preventDefault();
      wrapSelection(ta, '**', '**', '加粗文字');
      return;
    }
    if ((e.metaKey || e.ctrlKey) && (e.key === 'i' || e.key === 'I')) {
      e.preventDefault();
      wrapSelection(ta, '*', '*', '斜体文字');
      return;
    }

    if (e.key === 'Tab') {
      e.preventDefault();
      const start = ta.selectionStart;
      const end = ta.selectionEnd;
      const next = ta.value.slice(0, start) + '  ' + ta.value.slice(end);
      onChange(next);
      requestAnimationFrame(() => {
        ta.setSelectionRange(start + 2, start + 2);
      });
    }
  }

  function wrapSelection(ta: HTMLTextAreaElement, before: string, after: string, placeholder: string) {
    const start = ta.selectionStart;
    const end = ta.selectionEnd;
    const selected = ta.value.slice(start, end) || placeholder;
    const next = ta.value.slice(0, start) + before + selected + after + ta.value.slice(end);
    onChange(next);
    const cursorStart = start + before.length;
    const cursorEnd = cursorStart + selected.length;
    requestAnimationFrame(() => {
      ta.focus();
      ta.setSelectionRange(cursorStart, cursorEnd);
    });
  }

  return (
    <textarea
      ref={textareaRef}
      className="editor-textarea"
      spellCheck={false}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      onKeyDown={onKeyDown}
      placeholder="在这里输入 Markdown..."
    />
  );
}
