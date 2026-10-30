import { API } from "../constants/app";
import type { EngineInfo, HistoryRecord, LineageInspect, ProgressEvent, TaskDetail, TemplateInfo } from "../types";

export class ApiError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ApiError";
  }
}

async function readError(response: Response): Promise<string> {
  try {
    const body = (await response.json()) as { message?: string };
    return body.message || `请求失败（${response.status}）`;
  } catch {
    return `请求失败（${response.status}）`;
  }
}

export async function fetchEngine(): Promise<EngineInfo> {
  const response = await fetch(API.engines);
  if (!response.ok) {
    throw new ApiError(await readError(response));
  }
  return response.json() as Promise<EngineInfo>;
}

export async function fetchTemplates(): Promise<TemplateInfo[]> {
  const response = await fetch(API.templates);
  if (!response.ok) {
    throw new ApiError(await readError(response));
  }
  return response.json() as Promise<TemplateInfo[]>;
}

export async function fetchHistoryByHash(hash: string): Promise<HistoryRecord[]> {
  const response = await fetch(API.historyByHash(hash));
  if (!response.ok) {
    throw new ApiError(await readError(response));
  }
  const body = (await response.json()) as { items: HistoryRecord[] };
  return body.items;
}

export async function inspectLineage(file: File): Promise<LineageInspect> {
  const form = new FormData();
  form.append("file", file);
  const response = await fetch(API.historyInspect, { method: "POST", body: form });
  if (!response.ok) {
    throw new ApiError(await readError(response));
  }
  return response.json() as Promise<LineageInspect>;
}

export async function submitConvert(form: FormData): Promise<{ taskId: string }> {
  const response = await fetch(API.convert, { method: "POST", body: form });
  if (!response.ok) {
    throw new ApiError(await readError(response));
  }
  const body = (await response.json()) as { task_id: string };
  return { taskId: body.task_id };
}

export async function fetchTask(taskId: string): Promise<TaskDetail> {
  const response = await fetch(API.task(taskId));
  if (!response.ok) {
    throw new ApiError(await readError(response));
  }
  return response.json() as Promise<TaskDetail>;
}

/** 订阅任务 SSE。空日志是心跳，调用方不用处理。 */
export async function subscribeTask(
  taskId: string,
  onEvent: (event: ProgressEvent) => void,
  signal: AbortSignal,
): Promise<void> {
  const response = await fetch(API.taskEvents(taskId), {
    signal,
    headers: { Accept: "text/event-stream" },
  });
  if (!response.ok || !response.body) {
    throw new ApiError(await readError(response));
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      return;
    }
    buffer += decoder.decode(value, { stream: true });
    const chunks = buffer.split("\n\n");
    buffer = chunks.pop() ?? "";
    for (const chunk of chunks) {
      const line = chunk.split("\n").find((item) => item.startsWith("data:"));
      if (!line) {
        continue;
      }
      const event = JSON.parse(line.slice(5).trim()) as ProgressEvent;
      if (event.type === "log" && !event.message) {
        continue;
      }
      onEvent(event);
    }
  }
}
