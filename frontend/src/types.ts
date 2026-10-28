export interface EngineInfo {
  name: string;
  available: boolean;
  refresh_fields: boolean;
  notes: string[];
}

export interface TemplateInfo {
  template_id: string;
  name: string;
  kind: string;
  filename: string;
  is_default: boolean;
}

export interface HistoryRecord {
  record_id: string;
  created_at: string;
  source_name: string;
  output_name: string | null;
  success: boolean;
  message: string;
  duration_ms: number;
  warnings: string[];
}

export interface ProgressEvent {
  seq: number;
  type: "progress" | "log" | "done" | "error";
  step: string | null;
  index: number;
  total: number;
  percent: number;
  message: string;
  data?: { download_url?: string; output_name?: string; code?: string } | null;
}

export interface ValidationIssue {
  severity: "error" | "warning" | "off" | string;
  code: string;
  message: string;
  location?: string;
  excerpt?: string;
}

export interface TaskDetail {
  task_id: string;
  state: string;
  warnings: string[];
  message: string;
  output_name: string | null;
  download_url: string | null;
  error: { code?: string; message?: string; issues?: ValidationIssue[] } | null;
}
