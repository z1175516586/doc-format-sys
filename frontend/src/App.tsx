import { useEffect, useMemo, useState } from "react";
import {
  Alert,
  Button,
  Card,
  Radio,
  Col,
  Empty,
  Layout,
  List,
  Progress,
  Row,
  Select,
  Space,
  Tag,
  Tree,
  Typography,
  Upload,
  message,
} from "antd";
import type { UploadFile } from "antd";
import {
  CircleAlert,
  CircleCheck,
  Download,
  FileText,
  GitBranch,
  History,
  Info,
  Play,
} from "lucide";

import {
  fetchEngine,
  fetchTask,
  fetchTemplates,
  inspectLineage,
  submitConvert,
  subscribeTask,
} from "./api/client";
import { AppIcon } from "./components/AppIcon";
import {
  ACCEPTED_EXTENSIONS,
  API,
  APP_TITLE,
  DEFAULT_TOC_PAGE_SOURCE,
  DOCUMENT_TYPES,
  LINEAGE,
  NOTICES,
  STEP_LABELS,
  TOC_PAGE_SOURCE,
  TOC_PAGE_SOURCE_OPTIONS,
  TOC_TITLE,
} from "./constants/app";
import type { TocPageSource } from "./constants/app";
import type { EngineInfo, LineageInspect, LineageNode, TemplateInfo, ValidationIssue } from "./types";
import type { DataNode } from "antd/es/tree";

const { Header, Content } = Layout;
const { Title, Paragraph, Text } = Typography;

function lineageKey(node: LineageNode): string {
  return node.node_id || node.record_id || node.name;
}

function lineageTreeData(node: LineageNode): DataNode {
  const failed = node.kind === "failed";
  const title = (
    <Space size={4} wrap>
      <AppIcon icon={failed ? CircleAlert : FileText} size={14} />
      <Text strong={node.current} type={failed ? "danger" : undefined}>
        {failed ? `${LINEAGE.failed}${node.message ? `：${node.message}` : ""}` : node.name}
      </Text>
      {node.current ? <Tag color="blue">当前</Tag> : null}
      {failed ? <Tag color="error">失败</Tag> : null}
      {node.file_available && node.download_url ? (
        <a href={node.download_url} onClick={(event) => event.stopPropagation()}>
          <Space size={2}>
            <AppIcon icon={Download} size={12} />
            下载
          </Space>
        </a>
      ) : null}
      {!failed && !node.file_available ? <Text type="secondary">{node.missing_file_message || LINEAGE.missingFile}</Text> : null}
    </Space>
  );
  return {
    key: lineageKey(node),
    title,
    children: node.children.map(lineageTreeData),
  };
}

export default function App() {
  const [engine, setEngine] = useState<EngineInfo | null>(null);
  const [templates, setTemplates] = useState<TemplateInfo[]>([]);
  const [documentType, setDocumentType] = useState<number | undefined>();
  const [templateId, setTemplateId] = useState<string | undefined>();
  const [file, setFile] = useState<File | null>(null);
  const [fileList, setFileList] = useState<UploadFile[]>([]);
  const [lineage, setLineage] = useState<LineageInspect | null>(null);
  const [historyError, setHistoryError] = useState("");
  const [running, setRunning] = useState(false);
  const [percent, setPercent] = useState(0);
  const [stepText, setStepText] = useState("");
  const [downloadUrl, setDownloadUrl] = useState("");
  const [outputName, setOutputName] = useState("");
  const [warnings, setWarnings] = useState<string[]>([]);
  const [issues, setIssues] = useState<ValidationIssue[]>([]);
  /**
   * 目录页码由谁计算。二选一：这两种是同一件事的两种实现——
   * 服务端算完落盘，客户端打开时又会重算一遍，同时开只是白占全局单点的引擎。
   */
  const [pageSource, setPageSource] = useState<TocPageSource>(DEFAULT_TOC_PAGE_SOURCE);
  /** LibreOffice 被其他任务占用时，后端会推 waiting，这里切成「等待中」 */
  const [waiting, setWaiting] = useState(false);
  const [waitingMessage, setWaitingMessage] = useState("");

  useEffect(() => {
    fetchEngine().then(setEngine).catch(() => setEngine(null));
    fetchTemplates().then(setTemplates).catch(() => setTemplates([]));
  }, []);

  const contentTemplates = useMemo(
    () => templates.filter((item) => item.kind === "content"),
    [templates],
  );

  /** 服务端刷新依赖 LibreOffice，检测不到就禁用该选项，不能让用户选了必失败。 */
  const serverUnavailable = engine ? !engine.available || !engine.refresh_fields : false;

  // 引擎不可用时把已选的服务端方案回落到客户端，避免提交一个注定降级的任务
  useEffect(() => {
    if (serverUnavailable) {
      setPageSource((current) => (current === "server" ? DEFAULT_TOC_PAGE_SOURCE : current));
    }
  }, [serverUnavailable]);

  async function loadHistory(next: File) {
    setLineage(null);
    setHistoryError("");
    try {
      setLineage(await inspectLineage(next));
    } catch (error) {
      setHistoryError(error instanceof Error ? error.message : "查询转换树失败");
    }
  }

  function resetResult() {
    setPercent(0);
    setStepText("");
    setDownloadUrl("");
    setOutputName("");
    setWarnings([]);
    setIssues([]);
    setWaiting(false);
    setWaitingMessage("");
  }

  async function onConvert() {
    if (!file) {
      message.error("请先选择要转换的文档文件");
      return;
    }
    if (!documentType) {
      message.error("请选择一个排版模板");
      return;
    }

    const form = new FormData();
    form.append("file", file);
    form.append("document_type", String(documentType));
    form.append("toc_title", TOC_TITLE);
    form.append("header_text", "");
    // 互斥：服务端刷新与客户端更新只能二选一，同时为真会让引擎白算一遍
    form.append("update_fields_on_open", String(pageSource === "client"));
    form.append("server_refresh_fields", String(pageSource === "server"));
    if (templateId) {
      form.append("template_id", templateId);
    }

    resetResult();
    setRunning(true);
    const controller = new AbortController();

    try {
      const { taskId } = await submitConvert(form);
      await subscribeTask(
        taskId,
        (event) => {
          if (event.type === "progress") {
            setPercent(event.percent);
            // 引擎被占用时会推 engine_phase=waiting，此时明确显示排队，而不是干等
            const isWaiting = event.data?.waiting === true;
            setWaiting(isWaiting);
            if (isWaiting) {
              // 按钮上用短标签，长句放提示里
              setWaitingMessage(event.message || TOC_PAGE_SOURCE.waitingHint);
              setStepText(TOC_PAGE_SOURCE.waitingBadge);
              return;
            }
            setWaitingMessage("");
            const label = event.step ? STEP_LABELS[event.step] ?? event.step : "";
            setStepText(label ? `${label}（${event.index}/${event.total}）` : event.message);
          }
          if (event.type === "error") {
            setWaiting(false);
            message.error(event.message || "转换失败");
          }
          if (event.type === "done") {
            setPercent(100);
            setWaiting(false);
            setStepText("排版完成");
            setDownloadUrl(event.data?.download_url || API.taskDownload(taskId));
            setOutputName(event.data?.output_name || "排版结果.docx");
          }
        },
        controller.signal,
      );

      const detail = await fetchTask(taskId);
      setWarnings(detail.warnings ?? []);
      setIssues(detail.error?.issues ?? []);
      if (detail.state === "failed") {
        message.error(detail.error?.message || detail.message || "转换失败");
        setDownloadUrl("");
      } else if (detail.download_url) {
        setDownloadUrl(detail.download_url);
        setOutputName(detail.output_name || "排版结果.docx");
      }
      await loadHistory(file);
    } catch (error) {
      if (!(error instanceof DOMException && error.name === "AbortError")) {
        message.error(error instanceof Error ? error.message : "请求失败");
      }
    } finally {
      setRunning(false);
    }
  }

  const actionIcon = downloadUrl ? CircleCheck : running ? Play : FileText;

  return (
    <Layout style={{ minHeight: "100vh", background: "#f3f5f8" }}>
      <Header
        style={{
          background: "#14324f",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          paddingInline: 24,
        }}
      >
        <Space>
          <AppIcon icon={FileText} size={22} label={APP_TITLE} />
          <Title level={4} style={{ color: "#fff", margin: 0 }}>
            {APP_TITLE}
          </Title>
        </Space>
        {engine ? (
          <Tag color={engine.refresh_fields ? "green" : "gold"}>
            {engine.refresh_fields ? "目录页码可自动刷新" : "目录页码需在 Word 中按 F9 更新"}
          </Tag>
        ) : null}
      </Header>

      <Content style={{ padding: 24, maxWidth: 1280, margin: "0 auto", width: "100%" }}>
        <Row gutter={[16, 16]}>
          <Col xs={24} lg={6}>
            <Card
              title={
                <Space>
                  <AppIcon icon={Info} />
                  注意事项
                </Space>
              }
            >
              <List
                size="small"
                dataSource={[...NOTICES]}
                renderItem={(item) => <List.Item style={{ paddingInline: 0 }}>{item}</List.Item>}
              />
            </Card>
          </Col>

          <Col xs={24} lg={12}>
            <Card>
              <Title level={3} style={{ marginTop: 0 }}>
                <Space>
                  <AppIcon icon={actionIcon} size={22} />
                  上传文档并排版
                </Space>
              </Title>
              <Paragraph type="secondary">上传原文档，选择文档类型，完成后下载排版结果。</Paragraph>

              <Upload.Dragger
                accept={ACCEPTED_EXTENSIONS.join(",")}
                maxCount={1}
                fileList={fileList}
                beforeUpload={(next) => {
                  const name = next.name.toLowerCase();
                  const allowed = ACCEPTED_EXTENSIONS.some((ext) => name.endsWith(ext));
                  if (!allowed) {
                    message.error("仅支持 .doc 与 .docx");
                    return Upload.LIST_IGNORE;
                  }
                  setFile(next);
                  setFileList([{ uid: next.uid, name: next.name, status: "done" }]);
                  resetResult();
                  void loadHistory(next);
                  return false;
                }}
                onRemove={() => {
                  setFile(null);
                  setFileList([]);
                  setLineage(null);
                  setHistoryError("");
                  resetResult();
                }}
              >
                <Space direction="vertical">
                  <AppIcon icon={FileText} size={28} />
                  <Text>点击或拖入 Word 文档（.doc / .docx）</Text>
                </Space>
              </Upload.Dragger>

              <Space direction="vertical" size="middle" style={{ width: "100%", marginTop: 16 }}>
                <Select
                  style={{ width: "100%" }}
                  placeholder="选择排版模板"
                  value={documentType}
                  options={DOCUMENT_TYPES.map((item) => ({ value: item.value, label: item.label }))}
                  onChange={setDocumentType}
                />
                {contentTemplates.length > 0 ? (
                  <Select
                    allowClear
                    style={{ width: "100%" }}
                    placeholder="页眉页脚模板（可选，默认使用系统模板）"
                    value={templateId}
                    options={contentTemplates.map((item) => ({
                      value: item.template_id,
                      label: item.is_default ? `${item.name}（默认）` : item.name,
                    }))}
                    onChange={setTemplateId}
                  />
                ) : null}

                <div>
                  <Text strong>{TOC_PAGE_SOURCE.title}</Text>
                  <Radio.Group
                    style={{ display: "block", marginTop: 4 }}
                    value={pageSource}
                    onChange={(event) => setPageSource(event.target.value as TocPageSource)}
                  >
                    <Space direction="vertical" size={2} style={{ width: "100%" }}>
                      {TOC_PAGE_SOURCE_OPTIONS.map((option) => {
                        const disabled = option.value === "server" && serverUnavailable;
                        return (
                          <div key={option.value}>
                            <Radio value={option.value} disabled={disabled}>
                              {option.label}
                            </Radio>
                            <div>
                              <Text type="secondary" style={{ paddingInlineStart: 24, fontSize: 12 }}>
                                {disabled ? TOC_PAGE_SOURCE.engineUnavailable : option.hint}
                              </Text>
                            </div>
                          </div>
                        );
                      })}
                    </Space>
                  </Radio.Group>
                </div>

                <Button type="primary" block size="large" loading={running} onClick={() => void onConvert()}>
                  {running ? stepText || "转换中" : "开始转换文档"}
                </Button>
                {waiting ? (
                  <Alert
                    type="info"
                    showIcon
                    icon={<AppIcon icon={Info} />}
                    message={TOC_PAGE_SOURCE.waitingBadge}
                    description={waitingMessage || TOC_PAGE_SOURCE.waitingHint}
                  />
                ) : null}
                {running || percent > 0 ? <Progress percent={percent} status={downloadUrl ? "success" : "active"} /> : null}
              </Space>

              {issues.some((item) => item.severity === "error") ? (
                <Alert
                  style={{ marginTop: 16 }}
                  type="error"
                  showIcon
                  icon={<AppIcon icon={CircleAlert} />}
                  message="文档不符合排版规范，转换已终止"
                  description={
                    <Space direction="vertical" size={4}>
                      {issues.map((item, index) => (
                        <Text
                          key={`${item.code}-${index}`}
                          type={item.severity === "error" ? "danger" : "warning"}
                        >
                          {item.location ? `${item.location}：` : ""}
                          {item.message}
                        </Text>
                      ))}
                    </Space>
                  }
                />
              ) : null}

              {downloadUrl ? (
                <Alert
                  style={{ marginTop: 16 }}
                  type="success"
                  showIcon
                  icon={<AppIcon icon={CircleCheck} />}
                  message="转换成功"
                  description={
                    <Space direction="vertical">
                      <a href={downloadUrl} download={outputName}>
                        <Space>
                          <AppIcon icon={Download} />
                          下载转换后的文档
                        </Space>
                      </a>
                      {warnings.map((item) => (
                        <Text key={item} type="warning">
                          {item}
                        </Text>
                      ))}
                    </Space>
                  }
                />
              ) : null}
            </Card>
          </Col>

          <Col xs={24} lg={6}>
            <Card
              title={
                <Space>
                  <AppIcon icon={History} />
                  转换树
                </Space>
              }
              styles={{ body: { maxHeight: 640, overflow: "auto" } }}
            >
              {!file ? <Empty description={LINEAGE.empty} /> : null}
              {historyError ? <Alert type="error" showIcon message={historyError} /> : null}
              {file && lineage?.new_document ? <Empty description={LINEAGE.firstUpload} /> : null}
              {lineage?.modified ? (
                <Alert
                  style={{ marginBottom: 12 }}
                  type="warning"
                  showIcon
                  icon={<AppIcon icon={CircleAlert} />}
                  message={LINEAGE.modifiedTitle}
                  description={lineage.modified_message}
                />
              ) : null}
              {lineage?.path && lineage.path.length > 0 ? (
                <Space direction="vertical" size={12} style={{ width: "100%" }}>
                  <div>
                    <Text type="secondary">{LINEAGE.pathTitle}</Text>
                    <div>
                      {lineage.path.map((item, index) => (
                        <span key={lineageKey(item)}>
                          {index > 0 ? " → " : null}
                          <Text strong={item.current}>{item.name}</Text>
                        </span>
                      ))}
                    </div>
                  </div>
                  {lineage.tree ? (
                    <div>
                      <Space size={4} style={{ marginBottom: 8 }}>
                        <AppIcon icon={GitBranch} size={14} />
                        <Text type="secondary">{LINEAGE.treeTitle}</Text>
                      </Space>
                      <Tree
                        key={lineage.root_id || lineage.current_node_id || "tree"}
                        showLine
                        defaultExpandAll
                        selectable={false}
                        selectedKeys={lineage.current_node_id ? [lineage.current_node_id] : []}
                        treeData={[lineageTreeData(lineage.tree)]}
                        titleRender={(item) =>
                          typeof item.title === "function" ? item.title(item) : item.title
                        }
                      />
                    </div>
                  ) : null}
                </Space>
              ) : null}
            </Card>
          </Col>
        </Row>

        {engine && !engine.available ? (
          <Alert
            style={{ marginTop: 16 }}
            type="warning"
            showIcon
            icon={<AppIcon icon={CircleAlert} />}
            message="当前没有排版引擎，目录页码不会自动计算。"
          />
        ) : null}
      </Content>
    </Layout>
  );
}
