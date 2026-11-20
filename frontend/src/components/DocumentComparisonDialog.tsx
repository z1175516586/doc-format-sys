import { useEffect, useMemo, useRef, useState } from "react";
import { Alert, Button, Empty, Modal, Pagination, Segmented, Spin } from "antd";

import { fetchComparison } from "../api/client";
import type { ComparisonRow, DocumentComparison } from "../types";
import "./document-comparison.css";

const PAGE_SIZE = 80;
type Filter = "diff" | "all" | "style";

function ComparisonCell({ row, side }: { row: ComparisonRow; side: "before" | "after" }) {
  const value = row[side];
  const changed = row.kind !== "equal";
  const sign = changed ? (side === "before" ? "−" : "+") : "";

  return (
    <div className={"comparison-cell " + (!value ? "comparison-gap" : changed ? "comparison-" + side : "")}>
      <span className="comparison-gutter" aria-label={value ? "段落或表格行 " + value.number : undefined}>
        {value ? <><span>{sign}</span>{value.number}</> : "—"}
      </span>
      {value ? (
        <div className="comparison-cell-body">
          {row.block_type === "table_row" ? <span className="comparison-row-type">表格行</span> : null}
          <div className="comparison-text">
            {value.text.map((part, index) =>
              part.changed ? <mark key={index}>{part.text}</mark> : <span key={index}>{part.text}</span>,
            )}
          </div>
          {row.styles.length ? (
            <ul className="comparison-styles" aria-label={side === "before" ? "原样式" : "新样式"}>
              {row.styles.map((style) => (
                <li key={style.label}>
                  <span>{style.label}</span> {style[side]}{style.range_changed ? "（应用范围变化）" : ""}
                </li>
              ))}
            </ul>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}

export function DocumentComparisonDialog({ taskId, onClose }: { taskId: string; onClose: () => void }) {
  const [result, setResult] = useState<DocumentComparison | null>(null);
  const [error, setError] = useState("");
  const [attempt, setAttempt] = useState(0);
  const [filter, setFilter] = useState<Filter>("diff");
  const [page, setPage] = useState(1);
  const scroll = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const controller = new AbortController();
    fetchComparison(taskId, controller.signal)
      .then((data) => { if (!controller.signal.aborted) setResult(data); })
      .catch((reason: unknown) => {
        if (!controller.signal.aborted) {
          setError(reason instanceof Error ? reason.message : "读取文档对比失败");
        }
      });
    return () => controller.abort();
  }, [taskId, attempt]);

  const rows = useMemo(
    () => result?.rows.filter((row) =>
      filter === "all" || (filter === "style" ? row.styles.length > 0 : row.kind !== "equal"),
    ) ?? [],
    [result, filter],
  );

  function changePage(next: number) {
    setPage(next);
    scroll.current?.scrollTo({ top: 0 });
  }

  return (
    <Modal
      open
      onCancel={onClose}
      title="转换前后对比"
      width="min(1320px, calc(100vw - 28px))"
      style={{ top: 28 }}
      footer={<Button onClick={onClose}>关闭</Button>}
    >
      <p className="comparison-scope">
        对比正文、表格文字及常用段落和字符样式。页眉页脚、图片和页面布局请在 Word 中查看。
      </p>
      {error ? (
        <Alert
          type="error"
          showIcon
          title="无法加载对比"
          description={error}
          action={
            <Button size="small" onClick={() => { setError(""); setAttempt((value) => value + 1); }}>
              重试
            </Button>
          }
        />
      ) : result ? (
        <>
          <div className="comparison-toolbar">
            <Segmented<Filter>
              aria-label="筛选对比结果"
              value={filter}
              options={[
                { value: "diff", label: "仅差异" },
                { value: "all", label: "全部" },
                { value: "style", label: "样式变更" },
              ]}
              onChange={(value) => { setFilter(value); changePage(1); }}
            />
            <div className="comparison-counts">
              <span className="comparison-count-removed">− {result.counts.removed} 删除</span>
              <span className="comparison-count-added">+ {result.counts.added} 新增</span>
              <span>{result.counts.modified} 文字修改</span>
              <span>{result.rows.filter((row) => row.styles.length > 0).length} 样式变更</span>
            </div>
          </div>
          <div className="comparison-scroll" ref={scroll} tabIndex={0} aria-label="左右对比结果">
            <div className="comparison-pair comparison-heading">
              <div><strong>转换前</strong><span title={result.source_name}>{result.source_name}</span></div>
              <div><strong>转换后</strong><span title={result.output_name}>{result.output_name}</span></div>
            </div>
            {rows.length ? rows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE).map((row) => (
              <div className="comparison-pair" key={(row.before?.number ?? 0) + ":" + (row.after?.number ?? 0)}>
                <ComparisonCell row={row} side="before" />
                <ComparisonCell row={row} side="after" />
              </div>
            )) : <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="当前筛选下没有可显示的条目" />}
          </div>
          {rows.length > PAGE_SIZE ? (
            <Pagination
              className="comparison-pagination"
              size="small"
              current={page}
              onChange={changePage}
              pageSize={PAGE_SIZE}
              showSizeChanger={false}
              total={rows.length}
              showTotal={(total) => "共 " + total + " 行"}
            />
          ) : null}
        </>
      ) : <div className="comparison-loading"><Spin size="large" aria-label="正在对比文档" /></div>}
    </Modal>
  );
}
