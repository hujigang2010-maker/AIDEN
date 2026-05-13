export type StyleMap = Record<string, string>;

export interface Theme {
  id: string;
  name: string;
  description: string;
  base: StyleMap;
  styles: {
    container: StyleMap;
    h1: StyleMap;
    h2: StyleMap;
    h3: StyleMap;
    h4: StyleMap;
    h5: StyleMap;
    h6: StyleMap;
    p: StyleMap;
    strong: StyleMap;
    em: StyleMap;
    del: StyleMap;
    a: StyleMap;
    ul: StyleMap;
    ol: StyleMap;
    li: StyleMap;
    blockquote: StyleMap;
    blockquoteP: StyleMap;
    hr: StyleMap;
    code: StyleMap;
    pre: StyleMap;
    preCode: StyleMap;
    img: StyleMap;
    table: StyleMap;
    thead: StyleMap;
    th: StyleMap;
    tr: StyleMap;
    td: StyleMap;
    figure: StyleMap;
    figcaption: StyleMap;
  };
}
