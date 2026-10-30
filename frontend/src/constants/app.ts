/** 界面文案与业务枚举。改文案、改文档类型只动这个文件。 */

export const APP_TITLE = "高阶段智能排版";

export const TOC_TITLE = "目      录";

export const ACCEPTED_EXTENSIONS = [".doc", ".docx"] as const;

export const DOCUMENT_TYPES = [
  { value: 1, label: "初步设计" },
  { value: 2, label: "可行性研究" },
  { value: 3, label: "初步设计（代可研）" },
  { value: 4, label: "可行性研究（代初设）" },
] as const;

export const NOTICES = [
  "上传的文档至少包含四部分：封面、设计人员名单、目录和正文",
  "转换后某些空白字符（如空格、制表符）会有适当变化",
  "转换后表题和注释的样式会有适当变化",
  "标题编号必须是「1」「1.1」「1.2.1」这类数字与点交替的写法，且编号与标题之间留空格",
  "标题跳号、编号不规范或封面文档类型不对时，转换会终止并列出具体问题",
] as const;

/** 与后端 PIPELINE_STEPS 对应，供进度条展示。 */
export const STEP_LABELS: Record<string, string> = {
  open: "打开文档",
  precheck: "校验封面类型",
  repair: "修复结构",
  numbering: "编号转文本",
  conformance: "校验标题编号",
  orientation: "页面方向",
  footer: "清理页脚",
  highlight: "清除高亮",
  template: "套用页眉页脚",
  toc_cleanup: "清理旧目录",
  headings: "识别标题",
  split: "拆分封面与正文",
  cover: "生成封面",
  content: "整理正文",
  merge: "合并文档",
  toc_insert: "插入目录",
  toc_section: "目录分节",
  picture: "图片格式",
  finalize: "收尾",
  save: "保存",
  engine: "刷新目录页码",
  validate: "结构校验",
};

export const LINEAGE = {
  empty: "请上传文件以查看转换树",
  firstUpload: "该文件是首次进入本系统，还没有转换树",
  modifiedTitle: "文档内容已改动",
  missingFile: "有记录、文件已删",
  failed: "转换失败",
  pathTitle: "来时的路",
  treeTitle: "转换树",
} as const;

export const API = {
  convert: "/api/v1/documents/convert",
  engines: "/api/v1/engines",
  templates: "/api/v1/templates",
  historyInspect: "/api/v1/history/inspect",
  historyByHash: (hash: string) => `/api/v1/history/by-hash/${hash}`,
  historyFile: (recordId: string) => `/api/v1/history/${recordId}/file`,
  historyNodeFile: (nodeId: string) => `/api/v1/history/nodes/${nodeId}/file`,
  taskEvents: (taskId: string) => `/api/v1/tasks/${taskId}/events`,
  task: (taskId: string) => `/api/v1/tasks/${taskId}`,
  taskDownload: (taskId: string) => `/api/v1/tasks/${taskId}/download`,
} as const;
