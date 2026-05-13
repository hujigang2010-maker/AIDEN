import { forwardRef, useMemo } from 'react';
import DOMPurify from 'dompurify';

interface PreviewProps {
  fullHtml: string;
}

export const Preview = forwardRef<HTMLDivElement, PreviewProps>(function Preview(
  { fullHtml },
  ref,
) {
  const safeHtml = useMemo(() => {
    return DOMPurify.sanitize(fullHtml, {
      USE_PROFILES: { html: true },
      ADD_ATTR: ['target', 'style'],
    });
  }, [fullHtml]);

  return (
    <div className="preview-wrap">
      <div className="preview-frame">
        <div className="preview-bar">
          <span className="preview-dot" style={{ background: '#ff5f57' }} />
          <span className="preview-dot" style={{ background: '#febc2e' }} />
          <span className="preview-dot" style={{ background: '#28c840' }} />
          <span className="preview-title">公众号预览</span>
        </div>
        <div
          ref={ref}
          className="preview-content"
          dangerouslySetInnerHTML={{ __html: safeHtml }}
        />
      </div>
    </div>
  );
});
