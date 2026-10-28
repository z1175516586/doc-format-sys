import { useEffect, useMemo, useState } from "react";
import {
  Alert,
  Button,
  Card,
  Col,
  Empty,
  Layout,
  List,
  Progress,
  Row,
  Select,
  Space,
  Tag,
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
  History,
  Info,
  Play,
} from "lucide";

import {
  fetchEngine,
  fetchHistoryByHash,
  fetchTask,
  fetchTemplates,
  sha256OfFile,
  submitConvert,
  subscribeTask,
} from "./api/client";
import { AppIcon } from "./components/AppIcon";
import { ACCEPTED_EXTENSIONS, API, APP_TITLE, DOCUMENT_TYPES, NOTICES, STEP_LABELS, TOC_TITLE } from "./constants/app";
import type { EngineInfo, HistoryRecord, TemplateInfo } from "./types";

const { Header, Content } = Layout;
const { Title, Paragraph, Text } = Typography;

function formatTime(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return date.toLocaleString("zh-CN", { hour12: false });
}

export default function App() {
  const [engine, setEngine] = useState<EngineInfo | null>(null);
  const [templates, setTemplates] = useState<TemplateInfo[]>([]);
  const [documentType, setDocumentType] = useState<number | undefined>();
  const [templateId, setTemplateId] = useState<string | undefined>();
  const [file, setFile] = useState<File | null>(null);
  const [fileList, setFileList] = useState<UploadFile[]>([]);
  const [history, setHistory] = useState<HistoryRecord[] | null>(null);
  const [historyError, setHistoryError] = useState("");
  const [running, setRunning] = useState(false);
  const [percent, setPercent] = useState(0);
  const [stepText, setStepText] = useState("");
  const [downloadUrl, setDownloadUrl] = useState("");
  const [outputName, setOutputName] = useState("");
  const [warnings, setWarnings] = useState<string[]>([]);

  useEffect(() => {
    fetchEngine().then(setEngine).catch(() => setEngine(null));
    fetchTemplates().then(setTemplates).catch(() => setTemplates([]));
  }, []);

  const contentTemplates = useMemo(
    () => templates.filter((item) => item.kind === "content"),
    [templates],
  );

  async function loadHistory(next: File) {
    setHistory(null);
    setHistoryError("");
    try {
      const hash = await sha256OfFile(next);
      const items = await fetchHistoryByHash(hash);
      setHistory([...items].reverse());
    } catch (error) {
      setHistoryError(error instanceof Error ? error.message : "查询历史失败");
    }
  }

  function resetResult() {
    setPercent(0);
    setStepText("");
    setDownloadUrl("");
    setOutputName("");
    setWarnings([]);
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
            const label = event.step ? STEP_LABELS[event.step] ?? event.step : "";
            setStepText(label ? `${label}（${event.index}/${event.total}）` : event.message);
          }
          if (event.type === "error") {
            message.error(event.message || "转换失败");
          }
          if (event.type === "done") {
            setPercent(100);
            setStepText("排版完成");
            setDownloadUrl(event.data?.download_url || API.taskDownload(taskId));
            setOutputName(event.data?.output_name || "排版结果.docx");
          }
        },
        controller.signal,
      );

      const detail = await fetchTask(taskId);
      setWarnings(detail.warnings ?? []);
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
                  setHistory(null);
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
                <Button type="primary" block size="large" loading={running} onClick={() => void onConvert()}>
                  {running ? stepText || "转换中" : "开始转换文档"}
                </Button>
                {running || percent > 0 ? <Progress percent={percent} status={downloadUrl ? "success" : "active"} /> : null}
              </Space>

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
                  转换历史
                </Space>
              }
              styles={{ body: { maxHeight: 640, overflow: "auto" } }}
            >
              {!file ? <Empty description="请上传文件以查看历史记录" /> : null}
              {historyError ? <Alert type="error" showIcon message={historyError} /> : null}
              {file && history && history.length === 0 ? (
                <Empty description="该文件是首次上传，无转换历史" />
              ) : null}
              {history && history.length > 0 ? (
                <List
                  header={<Text>总转换次数：{history.length}</Text>}
                  dataSource={history}
                  renderItem={(record, index) => (
                    <List.Item>
                      <Space direction="vertical" size={2}>
                        <Text strong>第 {history.length - index} 次转换</Text>
                        <Text>
                          状态：
                          <Tag color={record.success ? "success" : "error"}>{record.success ? "成功" : "失败"}</Tag>
                        </Text>
                        <Text type="secondary">{formatTime(record.created_at)}</Text>
                        <Text>说明：{record.message || "无"}</Text>
                        {record.success && record.output_name ? (
                          <a href={API.historyFile(record.record_id)}>
                            <Space size={4}>
                              <AppIcon icon={Download} size={14} />
                              {record.output_name}
                            </Space>
                          </a>
                        ) : (
                          <Text type="secondary">下载文件：无</Text>
                        )}
                      </Space>
                    </List.Item>
                  )}
                />
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
