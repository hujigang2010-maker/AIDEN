import { useMemo, useState } from "react";
import { META, RESULTS, TREE, WEEKS, type TreeNode } from "./data";

function collectIds(node: TreeNode, into: Set<string>, depth = 0) {
  if (depth <= 1) into.add(node.id);
  node.children?.forEach((c) => collectIds(c, into, depth + 1));
}

function findNode(node: TreeNode, id: string): TreeNode | null {
  if (node.id === id) return node;
  for (const c of node.children ?? []) {
    const hit = findNode(c, id);
    if (hit) return hit;
  }
  return null;
}

export default function App() {
  const initialOpen = useMemo(() => {
    const s = new Set<string>();
    collectIds(TREE, s, 0);
    s.add("weeks");
    s.add("hospital");
    s.add("injury-care");
    s.add("stage");
    return s;
  }, []);
  const [open, setOpen] = useState<Set<string>>(initialOpen);
  const [selected, setSelected] = useState<string>("root");
  const [tab, setTab] = useState<"tree" | "weeks" | "end">("tree");
  const current = findNode(TREE, selected) ?? TREE;

  const toggle = (id: string) => {
    setOpen((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const expandAll = () => {
    const s = new Set<string>();
    const walk = (n: TreeNode) => {
      s.add(n.id);
      n.children?.forEach(walk);
    };
    walk(TREE);
    setOpen(s);
  };

  return (
    <div className="page">
      <header className="hero">
        <p className="eyebrow">内部处理总览 · 点击树枝展开</p>
        <h1>{META.title}</h1>
        <p className="sub">{META.subtitle}</p>
        <p className="meta">{META.date}</p>
        <ul className="pins">
          {META.pins.map((p) => (
            <li key={p}>{p}</li>
          ))}
        </ul>
      </header>

      <nav className="tabs" aria-label="主要板块">
        <button className={tab === "tree" ? "on" : ""} onClick={() => setTab("tree")}>
          1. 思维树
        </button>
        <button className={tab === "weeks" ? "on" : ""} onClick={() => setTab("weeks")}>
          2. 每周做什么
        </button>
        <button className={tab === "end" ? "on" : ""} onClick={() => setTab("end")}>
          3. 会得到什么
        </button>
      </nav>

      {tab === "tree" && (
        <section className="layout">
          <div className="tree-pane">
            <div className="pane-head">
              <h2>思维树</h2>
              <div className="tools">
                <button type="button" onClick={expandAll}>
                  全部展开
                </button>
                <button type="button" onClick={() => setOpen(new Set(["root"]))}>
                  收起到根
                </button>
              </div>
            </div>
            <p className="hint">先看根，再往下点。每点开一枝，右侧会显示：做什么、怎么推进、得到什么。</p>
            <ul className="tree">
              <TreeItem
                node={TREE}
                open={open}
                selected={selected}
                onToggle={toggle}
                onSelect={setSelected}
              />
            </ul>
          </div>
          <aside className="detail" aria-live="polite">
            <p className="kicker">{current.kicker ?? "这一枝"}</p>
            <h2>{current.title}</h2>
            <Triple card={current.card} />
          </aside>
        </section>
      )}

      {tab === "weeks" && (
        <section className="weeks">
          <p className="hint">时间从左到右。本周红框。每一项都拆成做、推、果。</p>
          <div className="week-grid">
            {WEEKS.map((w) => (
              <article key={w.id} className={`week-card phase-${w.phase}`}>
                <header>
                  <span className="week-label">{w.label}</span>
                  <span className="week-dates">{w.dates}</span>
                </header>
                <p className="week-goal">目标：{w.goal}</p>
                {w.items.map((it) => (
                  <div key={it.title} className="week-item">
                    <h3>{it.title}</h3>
                    <p className="owner">{it.owner}</p>
                    <Triple card={it.card} compact />
                  </div>
                ))}
              </article>
            ))}
          </div>
        </section>
      )}

      {tab === "end" && (
        <section className="ending">
          <p className="hint">过程结果可以本周勾；钱的结果要等认定书和评残，不要现在对外报价。</p>
          <div className="result-grid">
            {RESULTS.map((r) => (
              <article key={r.id} className="result-card">
                <h2>{r.title}</h2>
                <ul>
                  {r.items.map((x) => (
                    <li key={x}>{x}</li>
                  ))}
                </ul>
              </article>
            ))}
          </div>
          <div className="schemes">
            <article>
              <h3>方案甲 · 推荐结局</h3>
              <p>住院打通 → 保险进群 → 简易认定（同等或对方主责）→ 按发票赔已发生费用 → 三个月后评残补残疾赔偿金。钱主要来自保险公司。</p>
            </article>
            <article>
              <h3>方案乙 · 医院分叉</h3>
              <p>保险不认人民医院，或发热/伤口/外固定异常：由保险对接回齐鲁或指定三甲，其余同甲。不自己另找三甲。</p>
            </article>
            <article>
              <h3>方案丙 · 法律路径</h3>
              <p>失联或拒赔：律师函 → 认定书 → 起诉骑手（视情况平台）→ 执行盯保险。周期长，骑手个人财产不要指望。</p>
            </article>
          </div>
        </section>
      )}

      <footer>
        本文是家属内部梳理，不是医学鉴定、不是司法鉴定、不是律师函。对外以医院已审核报告、发票原件、交警认定书为准。
      </footer>
    </div>
  );
}

function Triple({
  card,
  compact,
}: {
  card: { do: string; how: string; result: string; fail?: string };
  compact?: boolean;
}) {
  return (
    <div className={`triple ${compact ? "compact" : ""}`}>
      <div>
        <h4>做什么</h4>
        <p>{card.do}</p>
      </div>
      <div>
        <h4>怎么推进</h4>
        <p>{card.how}</p>
      </div>
      <div>
        <h4>会得到什么</h4>
        <p>{card.result}</p>
      </div>
      {card.fail ? (
        <div className="fail">
          <h4>过不了就</h4>
          <p>{card.fail}</p>
        </div>
      ) : null}
    </div>
  );
}

function TreeItem({
  node,
  open,
  selected,
  onToggle,
  onSelect,
}: {
  node: TreeNode;
  open: Set<string>;
  selected: string;
  onToggle: (id: string) => void;
  onSelect: (id: string) => void;
}) {
  const hasKids = Boolean(node.children?.length);
  const expanded = open.has(node.id);
  return (
    <li>
      <div className={`node tone-${node.tone ?? "plain"} ${selected === node.id ? "sel" : ""}`}>
        {hasKids ? (
          <button
            type="button"
            className="twist"
            aria-label={expanded ? "收起" : "展开"}
            onClick={() => onToggle(node.id)}
          >
            {expanded ? "−" : "+"}
          </button>
        ) : (
          <span className="twist ghost">·</span>
        )}
        <button type="button" className="label" onClick={() => onSelect(node.id)}>
          {node.kicker ? <span className="tag">{node.kicker}</span> : null}
          {node.title}
        </button>
      </div>
      {hasKids && expanded ? (
        <ul>
          {node.children!.map((c) => (
            <TreeItem
              key={c.id}
              node={c}
              open={open}
              selected={selected}
              onToggle={onToggle}
              onSelect={onSelect}
            />
          ))}
        </ul>
      ) : null}
    </li>
  );
}
