const agentBase = "/api";
const SESSION_STORAGE_KEY = "office-agent.sessions.v1";
const FIRST_RESPONSE_TIMEOUT_MS = 15_000;

type ChatMessage = {
  role: "system" | "user" | "assistant";
  content: string;
};

type SmartInsertMode = "replace_selection" | "after_heading" | "append_end" | "insert_start";

type SmartInsertPlan = {
  mode: SmartInsertMode;
  anchorHeading?: string;
};

type InsertMode = "chat_only" | "replace_selection" | "smart_insert" | "append_end";

type ProviderConfigView = {
  baseUrl: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
  hasApiKey: boolean;
};

type ProviderModelsView = {
  ok: boolean;
  models: string[];
  currentModel?: string;
};

type KbStatsResponse = {
  ok: boolean;
  chunkCount: number;
  fileCount: number;
  files: string[];
};

type ChatSession = {
  id: string;
  title: string;
  updatedAt: number;
  messages: ChatMessage[];
  lastReply: string;
  lastUserInput: string;
};

type StreamDonePayload = {
  reply: string;
  retrievalCount?: number;
  citations?: Array<{ fileName: string }>;
};

type SmartInsertModelOutput = {
  mode?: SmartInsertMode;
  anchorHeading?: string;
  content?: string;
};

const $ = <T extends HTMLElement>(id: string): T => {
  const el = document.getElementById(id) as T | null;
  if (!el) {
    throw new Error(`Missing element: ${id}`);
  }
  return el;
};

const configStatus = $<HTMLParagraphElement>("configStatus");
const kbStatus = $<HTMLParagraphElement>("kbStatus");
const kbMeta = $<HTMLParagraphElement>("kbMeta");
const chatStatus = $<HTMLParagraphElement>("chatStatus");
const chatLog = $<HTMLDivElement>("chatLog");
const sendBtn = $<HTMLButtonElement>("sendMsg");
const sessionSelect = $<HTMLSelectElement>("sessionSelect");
const modelSelect = $<HTMLSelectElement>("model");

const state: {
  sessions: ChatSession[];
  currentSessionId: string;
  messages: ChatMessage[];
  lastReply: string;
  lastUserInput: string;
  isThinking: boolean;
  currentAbortController: AbortController | null;
} = {
  sessions: [],
  currentSessionId: "",
  messages: [],
  lastReply: "",
  lastUserInput: "",
  isThinking: false,
  currentAbortController: null,
};

function setStatus(target: HTMLParagraphElement, text: string): void {
  target.textContent = text;
}

function setThinking(isThinking: boolean): void {
  state.isThinking = isThinking;
  sendBtn.textContent = isThinking ? "停止" : "发送";
}

function appendMessage(role: "user" | "assistant", content: string): HTMLDivElement {
  const div = document.createElement("div");
  div.className = `msg ${role}`;
  div.textContent = `${role === "user" ? "你" : "助手"}: ${content}`;
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
  return div;
}

function renderChat(messages: ChatMessage[]): void {
  chatLog.innerHTML = "";
  for (const msg of messages) {
    if (msg.role === "system") {
      continue;
    }
    appendMessage(msg.role, msg.content);
  }
}

function cleanupMarkdownForWord(input: string): string {
  let text = input;

  text = text.replace(/```[\s\S]*?```/g, (block) => block.replace(/```[a-zA-Z0-9_-]*\n?|```/g, ""));
  text = text.replace(/^\s{0,3}#{1,6}\s+/gm, "");
  text = text.replace(/\*\*(.*?)\*\*/g, "$1");
  text = text.replace(/__(.*?)__/g, "$1");
  text = text.replace(/(^|[^*])\*(?!\s)([^*]+?)\*(?!\*)/g, "$1$2");
  text = text.replace(/(^|[^_])_(?!\s)([^_]+?)_(?!_)/g, "$1$2");
  text = text.replace(/`([^`]+)`/g, "$1");
  text = text.replace(/^\s*[-*+]\s+/gm, "- ");
  text = text.replace(/^\s*\d+\.\s+/gm, "");
  text = text.replace(/^>\s?/gm, "");
  text = text.replace(/\[([^\]]+)\]\(([^)]+)\)/g, "$1 ($2)");
  text = text.replace(/\n{3,}/g, "\n\n");

  return text.trim();
}

async function parseErrorMessage(response: Response): Promise<string> {
  const contentType = response.headers.get("content-type") || "";

  try {
    if (contentType.includes("application/json")) {
      const payload = (await response.json()) as { error?: string; message?: string };
      return payload.error || payload.message || JSON.stringify(payload);
    }
  } catch {
    return `HTTP ${response.status}`;
  }

  const text = await response.text();
  return text || `HTTP ${response.status}`;
}

function nowId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function buildSessionTitle(messages: ChatMessage[]): string {
  const firstUser = messages.find((msg) => msg.role === "user")?.content.trim() || "新会话";
  return firstUser.slice(0, 24) || "新会话";
}

function loadSessions(): void {
  try {
    const raw = localStorage.getItem(SESSION_STORAGE_KEY);
    const parsed = raw ? (JSON.parse(raw) as ChatSession[]) : [];
    state.sessions = parsed.filter((s) => s && s.id && Array.isArray(s.messages));
  } catch {
    state.sessions = [];
  }

  if (state.sessions.length === 0) {
    const first: ChatSession = {
      id: nowId(),
      title: "新会话",
      updatedAt: Date.now(),
      messages: [],
      lastReply: "",
      lastUserInput: "",
    };
    state.sessions.push(first);
  }

  state.currentSessionId = state.sessions[0].id;
  hydrateCurrentSession();
}

function saveSessions(): void {
  localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(state.sessions));
}

function hydrateCurrentSession(): void {
  const current = state.sessions.find((s) => s.id === state.currentSessionId);
  if (!current) {
    return;
  }
  state.messages = [...current.messages];
  state.lastReply = current.lastReply;
  state.lastUserInput = current.lastUserInput;
  renderSessionSelect();
  renderChat(state.messages);
}

function persistCurrentSession(): void {
  const idx = state.sessions.findIndex((s) => s.id === state.currentSessionId);
  if (idx < 0) {
    return;
  }

  const next: ChatSession = {
    ...state.sessions[idx],
    title: buildSessionTitle(state.messages),
    updatedAt: Date.now(),
    messages: [...state.messages],
    lastReply: state.lastReply,
    lastUserInput: state.lastUserInput,
  };

  state.sessions[idx] = next;
  state.sessions.sort((a, b) => b.updatedAt - a.updatedAt);
  saveSessions();
  renderSessionSelect();
}

function renderSessionSelect(): void {
  sessionSelect.innerHTML = "";
  for (const session of state.sessions) {
    const opt = document.createElement("option");
    opt.value = session.id;
    opt.textContent = session.title;
    sessionSelect.appendChild(opt);
  }
  sessionSelect.value = state.currentSessionId;
}

function createSession(): void {
  const session: ChatSession = {
    id: nowId(),
    title: "新会话",
    updatedAt: Date.now(),
    messages: [],
    lastReply: "",
    lastUserInput: "",
  };
  state.sessions.unshift(session);
  state.currentSessionId = session.id;
  hydrateCurrentSession();
  saveSessions();
  setStatus(chatStatus, "已创建新会话");
}

function deleteCurrentSession(): void {
  if (state.sessions.length <= 1) {
    setStatus(chatStatus, "至少保留一个会话");
    return;
  }
  state.sessions = state.sessions.filter((s) => s.id !== state.currentSessionId);
  state.currentSessionId = state.sessions[0].id;
  hydrateCurrentSession();
  saveSessions();
  setStatus(chatStatus, "已删除当前会话");
}

function clearCurrentSession(): void {
  state.messages = [];
  state.lastReply = "";
  state.lastUserInput = "";
  renderChat(state.messages);
  persistCurrentSession();
  setStatus(chatStatus, "已清空当前会话");
}

function upsertModelOptions(models: string[], preferred?: string): void {
  const merged = [...new Set(models.filter(Boolean))];
  if (preferred && !merged.includes(preferred)) {
    merged.unshift(preferred);
  }

  modelSelect.innerHTML = "";
  for (const model of merged) {
    const option = document.createElement("option");
    option.value = model;
    option.textContent = model;
    modelSelect.appendChild(option);
  }

  if (preferred) {
    modelSelect.value = preferred;
  }
}

async function refreshModels(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/provider/models`, { method: "GET" });
    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(configStatus, `读取模型失败: ${err}`);
      return;
    }

    const data = (await res.json()) as ProviderModelsView;
    upsertModelOptions(data.models || [], data.currentModel);
    setStatus(configStatus, `已加载 ${data.models.length} 个模型`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `读取模型失败: ${message}`);
  }
}

async function loadProviderConfig(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/provider/config`, { method: "GET" });

    if (!res.ok) {
      if (res.status === 404) {
        setStatus(configStatus, "尚未保存模型配置");
        return;
      }
      const err = await parseErrorMessage(res);
      setStatus(configStatus, `读取配置失败: ${err}`);
      return;
    }

    const data = (await res.json()) as ProviderConfigView;
    $<HTMLInputElement>("baseUrl").value = data.baseUrl || "";
    upsertModelOptions(data.model ? [data.model] : [], data.model);
    $<HTMLInputElement>("temperature").value = String(data.temperature ?? 0.2);
    $<HTMLInputElement>("maxTokens").value = String(data.maxTokens ?? 900);

    const apiKeyInput = $<HTMLInputElement>("apiKey");
    apiKeyInput.value = "";
    apiKeyInput.placeholder = data.hasApiKey ? "已保存（留空表示不修改）" : "sk-...";

    await refreshModels();
    setStatus(configStatus, data.hasApiKey ? "已加载已保存配置" : "已加载配置（尚未保存 API Key）");
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `读取配置失败: 无法连接本地 Agent (${message})`);
  }
}

async function saveProviderConfig(): Promise<void> {
  try {
    setStatus(configStatus, "保存中...");

    const model = modelSelect.value.trim();
    if (!model) {
      setStatus(configStatus, "请先刷新并选择模型");
      return;
    }

    const body = {
      baseUrl: $<HTMLInputElement>("baseUrl").value.trim(),
      apiKey: $<HTMLInputElement>("apiKey").value.trim(),
      model,
      temperature: Number($<HTMLInputElement>("temperature").value || "0.2"),
      maxTokens: Number($<HTMLInputElement>("maxTokens").value || "900"),
    };

    const res = await fetch(`${agentBase}/v1/provider/config`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(configStatus, `保存失败: ${err}`);
      return;
    }

    setStatus(configStatus, "保存成功（本地存储）");
    await refreshModels();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `保存失败: 无法连接本地 Agent (${message})`);
  }
}

async function loadKbStats(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/kb/stats`, { method: "GET" });
    if (!res.ok) {
      kbMeta.textContent = "知识库状态读取失败";
      return;
    }

    const data = (await res.json()) as KbStatsResponse;
    const latest = data.files.slice(-5).join("，");
    kbMeta.textContent = `知识库：${data.fileCount} 个文件 / ${data.chunkCount} 个分块${latest ? `。最近文件：${latest}` : ""}`;
  } catch {
    kbMeta.textContent = "知识库状态读取失败（本地 Agent 不可达）";
  }
}

async function uploadKnowledgeFile(): Promise<void> {
  const input = $<HTMLInputElement>("kbFile");
  if (!input.files || input.files.length === 0) {
    setStatus(kbStatus, "请先选择文件");
    return;
  }

  try {
    setStatus(kbStatus, "上传中...");

    const form = new FormData();
    form.append("file", input.files[0]);

    const res = await fetch(`${agentBase}/v1/kb/upload`, {
      method: "POST",
      body: form,
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(kbStatus, `上传失败: ${err}`);
      return;
    }

    const data = (await res.json()) as { chunkCount: number; fileName: string };
    setStatus(kbStatus, `已索引 ${data.fileName}，分块 ${data.chunkCount}`);
    await loadKbStats();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(kbStatus, `上传失败: 无法连接本地 Agent (${message})`);
  }
}

async function getWordContext(): Promise<{ documentContext: string; selection: string }> {
  if (!(window as any).Word || !(window as any).Office) {
    return { documentContext: "", selection: "" };
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const selection = context.document.getSelection();
    body.load("text");
    selection.load("text");
    await context.sync();

    return {
      documentContext: (body.text || "").slice(0, 3000),
      selection: selection.text || "",
    };
  });
}

function buildSmartInsertPlan(userInstruction: string, selectionText: string): SmartInsertPlan {
  const text = userInstruction.trim();
  const lower = text.toLowerCase();

  if (selectionText.trim()) {
    return { mode: "replace_selection" };
  }

  if (/\b(start|beginning|introduction|preface)\b/.test(lower) || /(开头|开始|前言|引言|前置)/.test(text)) {
    return { mode: "insert_start" };
  }

  if (/\b(append|at end|to end|conclusion|summary)\b/.test(lower) || /(结尾|末尾|最后|总结|结论)/.test(text)) {
    return { mode: "append_end" };
  }

  const headingMatch =
    text.match(/(?:在|到|于)\s*[“"']?([^”"'，。；\n]{2,40})[”"']?\s*(?:后|下面|之后|章节后)/) ||
    text.match(/(?:after|under)\s+["']?([^"'\n]{2,40})["']?/i);

  if (headingMatch?.[1]) {
    return { mode: "after_heading", anchorHeading: headingMatch[1].trim() };
  }

  return { mode: "append_end" };
}

async function applyTextToWordSelection(text: string): Promise<void> {
  if (!(window as any).Word) {
    throw new Error("当前不在 Word 宿主中");
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(text, "Replace");
    await context.sync();
  });
}

async function applyTextSmartly(text: string, plan: SmartInsertPlan): Promise<string> {
  if (!(window as any).Word) {
    throw new Error("当前不在 Word 宿主中");
  }

  const normalized = cleanupMarkdownForWord(text);
  if (!normalized) {
    throw new Error("没有可插入的正文内容");
  }

  return Word.run(async (context) => {
    const body = context.document.body;

    if (plan.mode === "replace_selection") {
      const selection = context.document.getSelection();
      selection.insertText(normalized, "Replace");
      await context.sync();
      return "已替换当前选区";
    }

    if (plan.mode === "insert_start") {
      body.insertText(`${normalized}\n`, "Start");
      await context.sync();
      return "已插入到文档开头";
    }

    if (plan.mode === "after_heading" && plan.anchorHeading) {
      const anchor = plan.anchorHeading.trim();
      const matches = body.search(anchor, { matchCase: false, matchWholeWord: false });
      matches.load("items");
      await context.sync();

      if (matches.items.length > 0) {
        matches.items[0].insertText(`\n${normalized}\n`, "After");
        await context.sync();
        return `已插入到“${anchor}”后`;
      }
    }

    body.insertText(`\n${normalized}\n`, "End");
    await context.sync();
    return "未找到锚点，已插入到文档末尾";
  });
}

async function sendMessageStream(
  payload: { messages: ChatMessage[]; documentContext: string; selection: string },
  signal: AbortSignal,
  onFirstChunk: () => void,
  onDelta: (text: string) => void
): Promise<StreamDonePayload> {
  const res = await fetch(`${agentBase}/v1/chat/stream`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    signal,
  });

  if (!res.ok) {
    const err = await parseErrorMessage(res);
    throw new Error(err);
  }

  if (!res.body) {
    throw new Error("流式响应不可用");
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let buffer = "";
  let fullReply = "";
  let donePayload: StreamDonePayload = { reply: "" };
  let firstChunkReceived = false;

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }

    if (!firstChunkReceived) {
      firstChunkReceived = true;
      onFirstChunk();
    }

    buffer += decoder.decode(value, { stream: true });
    const events = buffer.split("\n\n");
    buffer = events.pop() ?? "";

    for (const event of events) {
      const line = event
        .split("\n")
        .find((item) => item.trim().startsWith("data:"));
      if (!line) {
        continue;
      }

      const raw = line.slice(5).trim();
      if (!raw) {
        continue;
      }

      let data: { type?: string; delta?: string; reply?: string; retrievalCount?: number; citations?: Array<{ fileName: string }>; error?: string };
      try {
        data = JSON.parse(raw) as {
          type?: string;
          delta?: string;
          reply?: string;
          retrievalCount?: number;
          citations?: Array<{ fileName: string }>;
          error?: string;
        };
      } catch {
        continue;
      }

      if (data.type === "delta" && data.delta) {
        fullReply += data.delta;
        onDelta(fullReply);
      }

      if (data.type === "error") {
        throw new Error(data.error || "流式响应出错");
      }

      if (data.type === "done") {
        donePayload = {
          reply: data.reply ?? fullReply,
          retrievalCount: data.retrievalCount,
          citations: data.citations,
        };
      }
    }
  }

  return {
    reply: (donePayload.reply || fullReply).trim(),
    retrievalCount: donePayload.retrievalCount,
    citations: donePayload.citations,
  };
}

function getInsertMode(): InsertMode {
  return $<HTMLSelectElement>("insertMode").value as InsertMode;
}

function buildAgentPromptByMode(mode: InsertMode, text: string): string {
  if (mode === "chat_only") {
    return text;
  }

  if (mode === "smart_insert") {
    return `${text}\n\n你正在为 Word 文档执行“智能定位插入”。请阅读文档上下文后，仅返回 JSON 对象，不要返回其他文字：{"mode":"replace_selection|after_heading|append_end|insert_start","anchorHeading":"可选，字符串","content":"要写入 Word 的最终正文（纯文本，不含 Markdown）"}。如果拿不准位置，mode 使用 append_end。`;
  }

  return `${text}\n\n请只返回可直接写入 Word 的最终正文，不要使用 Markdown 标记（如 #、*、\`\`\`）。`;
}

function extractSmartInsertOutput(reply: string, instruction: string, selectionText: string): { plan: SmartInsertPlan; content: string } {
  const fallbackPlan = buildSmartInsertPlan(instruction, selectionText);
  const trimmed = reply.trim();

  const fenceMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const jsonText = (fenceMatch?.[1] || trimmed).trim();

  let parsed: SmartInsertModelOutput | null = null;
  try {
    parsed = JSON.parse(jsonText) as SmartInsertModelOutput;
  } catch {
    const objMatch = jsonText.match(/\{[\s\S]*\}/);
    if (objMatch) {
      try {
        parsed = JSON.parse(objMatch[0]) as SmartInsertModelOutput;
      } catch {
        parsed = null;
      }
    }
  }

  if (!parsed || typeof parsed !== "object") {
    return {
      plan: fallbackPlan,
      content: cleanupMarkdownForWord(reply),
    };
  }

  const mode = parsed.mode;
  const normalizedMode: SmartInsertMode =
    mode === "replace_selection" || mode === "after_heading" || mode === "append_end" || mode === "insert_start"
      ? mode
      : fallbackPlan.mode;

  const content = cleanupMarkdownForWord(parsed.content || "");
  return {
    plan: {
      mode: normalizedMode,
      anchorHeading: parsed.anchorHeading?.trim() || fallbackPlan.anchorHeading,
    },
    content: content || cleanupMarkdownForWord(reply),
  };
}

async function sendMessage(): Promise<void> {
  if (state.isThinking) {
    state.currentAbortController?.abort("user_stop");
    return;
  }

  const input = $<HTMLTextAreaElement>("userInput");
  const rawText = input.value.trim();
  if (!rawText) {
    setStatus(chatStatus, "请输入消息");
    return;
  }

  input.value = "";
  const mode = getInsertMode();
  const finalUserPrompt = buildAgentPromptByMode(mode, rawText);

  appendMessage("user", rawText);
  state.messages.push({ role: "user", content: finalUserPrompt });
  state.lastUserInput = rawText;
  persistCurrentSession();

  const assistantEl = appendMessage("assistant", "");
  const controller = new AbortController();
  let firstResponseTimer: ReturnType<typeof setTimeout> | null = null;
  state.currentAbortController = controller;
  setThinking(true);
  setStatus(chatStatus, "思考中（可点击停止）...");

  try {
    const wordContext = await getWordContext();

    firstResponseTimer = setTimeout(() => {
      if (!controller.signal.aborted) {
        controller.abort("first_response_timeout");
      }
    }, FIRST_RESPONSE_TIMEOUT_MS);

    const data = await sendMessageStream(
      {
        messages: state.messages,
        documentContext: wordContext.documentContext,
        selection: wordContext.selection,
      },
      controller.signal,
      () => {
        if (firstResponseTimer) {
          clearTimeout(firstResponseTimer);
          firstResponseTimer = null;
        }
      },
      (partial) => {
        assistantEl.textContent = `助手: ${partial}`;
      }
    );

    if (!data.reply) {
      throw new Error("未收到模型回复");
    }

    const citationText = data.citations?.length
      ? `\n\n来源: ${data.citations.map((c) => c.fileName).join(", ")}（命中 ${data.retrievalCount ?? data.citations.length} 段）`
      : "";

    if (mode === "smart_insert") {
      const smart = extractSmartInsertOutput(data.reply, rawText, wordContext.selection);
      const result = await applyTextSmartly(smart.content, smart.plan);
      assistantEl.textContent = `助手: ${smart.content}${citationText}`;
      state.lastReply = smart.content;
      state.messages.push({ role: "assistant", content: smart.content });
      persistCurrentSession();
      setStatus(chatStatus, `已完成并智能插入：${result}`);
      return;
    }

    assistantEl.textContent = `助手: ${data.reply}${citationText}`;
    state.lastReply = data.reply;
    state.messages.push({ role: "assistant", content: data.reply });
    persistCurrentSession();

    if (mode === "replace_selection") {
      await applyTextToWordSelection(cleanupMarkdownForWord(data.reply));
      setStatus(chatStatus, "已完成并替换当前选区");
      return;
    }

    if (mode === "append_end") {
      const result = await applyTextSmartly(data.reply, { mode: "append_end" });
      setStatus(chatStatus, `已完成并插入：${result}`);
      return;
    }

    setStatus(chatStatus, "已完成");
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    if (controller.signal.aborted) {
      const reason = String(controller.signal.reason || "");
      if (reason.includes("user_stop")) {
        setStatus(chatStatus, "已停止");
      } else if (reason.includes("first_response_timeout")) {
        setStatus(chatStatus, `请求超时：${FIRST_RESPONSE_TIMEOUT_MS / 1000} 秒内未收到首包`);
      } else {
        setStatus(chatStatus, "请求已取消或超时");
      }
    } else if (message.includes("timeout") || message.includes("超时")) {
      setStatus(chatStatus, `请求超时: ${message}`);
    } else {
      setStatus(chatStatus, `请求失败: ${message}`);
    }
  } finally {
    if (firstResponseTimer) {
      clearTimeout(firstResponseTimer);
    }
    state.currentAbortController = null;
    setThinking(false);
  }
}

async function insertLastReplyToWord(): Promise<void> {
  if (!state.lastReply) {
    setStatus(chatStatus, "没有可插入的回复");
    return;
  }

  const cleanText = cleanupMarkdownForWord(state.lastReply);
  await applyTextToWordSelection(cleanText);
  setStatus(chatStatus, "已替换当前选区");
}

async function insertLastReplyAtCursor(): Promise<void> {
  if (!state.lastReply) {
    setStatus(chatStatus, "没有可插入的回复");
    return;
  }

  if (!(window as any).Word) {
    setStatus(chatStatus, "当前不在 Word 宿主中");
    return;
  }

  const cleanText = cleanupMarkdownForWord(state.lastReply);
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(cleanText, "Start");
    await context.sync();
  });

  setStatus(chatStatus, "已将最后回复插入到光标");
}

async function retryLast(): Promise<void> {
  if (!state.lastUserInput) {
    setStatus(chatStatus, "没有可重试的用户消息");
    return;
  }
  $<HTMLTextAreaElement>("userInput").value = state.lastUserInput;
  await sendMessage();
}

function bindActions(): void {
  $("saveConfig").addEventListener("click", () => {
    void saveProviderConfig();
  });

  $("refreshModels").addEventListener("click", () => {
    void refreshModels();
  });

  $("uploadFile").addEventListener("click", () => {
    void uploadKnowledgeFile();
  });

  $("refreshKb").addEventListener("click", () => {
    void loadKbStats();
  });

  sendBtn.addEventListener("click", () => {
    void sendMessage();
  });

  $("insertReply").addEventListener("click", () => {
    void insertLastReplyToWord();
  });

  $("insertReplyCursor").addEventListener("click", () => {
    void insertLastReplyAtCursor();
  });

  $("retryLast").addEventListener("click", () => {
    void retryLast();
  });

  sessionSelect.addEventListener("change", () => {
    state.currentSessionId = sessionSelect.value;
    hydrateCurrentSession();
    setStatus(chatStatus, "已切换会话");
  });

  $("newSession").addEventListener("click", () => {
    createSession();
  });

  $("deleteSession").addEventListener("click", () => {
    deleteCurrentSession();
  });

  $("clearSession").addEventListener("click", () => {
    clearCurrentSession();
  });
}

Office.onReady(() => {
  bindActions();
  loadSessions();
  void loadProviderConfig();
  void loadKbStats();
  setThinking(false);
  setStatus(chatStatus, "就绪：可直接开始对话");
});
