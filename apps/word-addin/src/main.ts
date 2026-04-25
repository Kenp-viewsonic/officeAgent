const agentBase = "/api";

// ─── Types ───────────────────────────────────────────────────────────────────

type ChatMessage = {
  role: "system" | "user" | "assistant";
  content: string;
};

type DocumentParagraph = {
  index: number;
  text: string;
  style: string;
  headingLevel?: number;
  isTable: boolean;
  isList: boolean;
};

type DocumentStructure = {
  title: string;
  totalParagraphs: number;
  totalCharacters: number;
  paragraphs: DocumentParagraph[];
  selection: {
    text: string;
    startParagraphIndex?: number;
    endParagraphIndex?: number;
  };
};

type WordAction = {
  action: string;
  params: Record<string, any>;
  description: string;
};

type ActionPlan = {
  actions: WordAction[];
  explanation: string;
};

type Session = {
  id: string;
  title: string;
  messages: ChatMessage[];
  createdAt: number;
  updatedAt: number;
};

type InsertMode = "chat_only" | "smart_action" | "replace_selection" | "append_end";

type KbFileEntry = {
  fileName: string;
  chunkCount: number;
};

type SimpleInsertPlan = {
  text: string;
  mode: "replace" | "append_end" | "insert_start";
};

// ─── State ───────────────────────────────────────────────────────────────────

const STORAGE_KEY = "word-agent-sessions";

const state: {
  sessions: Session[];
  activeSessionId: string | null;
  lastReply: string;
  pendingActionPlan: ActionPlan | null;
  pendingSimpleInsert: SimpleInsertPlan | null;
} = {
  sessions: [],
  activeSessionId: null,
  lastReply: "",
  pendingActionPlan: null,
  pendingSimpleInsert: null,
};

// ─── DOM Helpers ─────────────────────────────────────────────────────────────

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
const kbFileList = $<HTMLDivElement>("kbFileList");
const chatStatus = $<HTMLParagraphElement>("chatStatus");
const chatLog = $<HTMLDivElement>("chatLog");
const sessionSelect = $<HTMLSelectElement>("sessionSelect");
const userInput = $<HTMLTextAreaElement>("userInput");
const insertModeSelect = $<HTMLSelectElement>("insertMode");
const actionPlanPanel = $<HTMLDivElement>("actionPlanPanel");
const actionPlanContent = $<HTMLDivElement>("actionPlanContent");

// ─── Provider Config ────────────────────────────────────────────────────────

type ProviderConfigView = {
  baseUrl: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
  firstTokenTimeout?: number;
  overallTimeout?: number;
  hasApiKey: boolean;
};

function setStatus(target: HTMLParagraphElement, text: string): void {
  target.textContent = text;
}

// ─── Chat Log ────────────────────────────────────────────────────────────────

function appendMessage(role: "user" | "assistant" | "system", content: string): void {
  const div = document.createElement("div");
  div.className = `msg ${role}`;
  div.textContent = `${role === "user" ? "你" : role === "assistant" ? "助手" : "系统"}: ${content}`;
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}

function createAssistantMessage(initial = ""): HTMLDivElement {
  const div = document.createElement("div");
  div.className = "msg assistant";
  div.textContent = `助手: ${initial}`;
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
  return div;
}

function updateAssistantMessage(el: HTMLDivElement, content: string): void {
  el.textContent = `助手: ${content}`;
  chatLog.scrollTop = chatLog.scrollHeight;
}

function clearChatLog(): void {
  chatLog.innerHTML = "";
}

// ─── Markdown Cleanup ───────────────────────────────────────────────────────

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

// ─── Error Parsing ───────────────────────────────────────────────────────────

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

// ─── Session Management ────────────────────────────────────────────────────

function generateSessionId(): string {
  return `session-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function loadSessionsFromStorage(): void {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      state.sessions = JSON.parse(raw) as Session[];
    }
  } catch {
    state.sessions = [];
  }
}

function saveSessionsToStorage(): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.sessions));
  } catch {
    // localStorage may be full or unavailable
  }
}

function getActiveSession(): Session | null {
  return state.sessions.find((s) => s.id === state.activeSessionId) ?? null;
}

function renderSessionSelect(): void {
  sessionSelect.innerHTML = "";
  for (const session of state.sessions) {
    const option = document.createElement("option");
    option.value = session.id;
    option.textContent = session.title || `会话 ${session.id.slice(8, 16)}`;
    if (session.id === state.activeSessionId) {
      option.selected = true;
    }
    sessionSelect.appendChild(option);
  }
}

function switchToSession(sessionId: string): void {
  // Save current session
  const current = getActiveSession();
  if (current) {
    current.updatedAt = Date.now();
  }
  saveSessionsToStorage();

  state.activeSessionId = sessionId;
  const target = getActiveSession();
  clearChatLog();

  if (target) {
    for (const msg of target.messages) {
      if (msg.role === "system") continue;
      appendMessage(msg.role, msg.content);
    }
    state.lastReply = [...target.messages].reverse().find((m) => m.role === "assistant")?.content ?? "";
  } else {
    state.lastReply = "";
  }

  renderSessionSelect();
  saveSessionsToStorage();
}

function createNewSession(): void {
  const session: Session = {
    id: generateSessionId(),
    title: "",
    messages: [],
    createdAt: Date.now(),
    updatedAt: Date.now(),
  };
  state.sessions.unshift(session);
  switchToSession(session.id);
  setStatus(chatStatus, "已创建新会话");
}

function clearCurrentSession(): void {
  const session = getActiveSession();
  if (!session) {
    setStatus(chatStatus, "没有当前会话");
    return;
  }
  session.messages = [];
  session.title = "";
  session.updatedAt = Date.now();
  clearChatLog();
  state.lastReply = "";
  saveSessionsToStorage();
  setStatus(chatStatus, "已清空当前会话");
}

function deleteCurrentSession(): void {
  if (state.sessions.length === 0) {
    setStatus(chatStatus, "没有可删除的会话");
    return;
  }
  const currentId = state.activeSessionId;
  state.sessions = state.sessions.filter((s) => s.id !== currentId);

  if (state.sessions.length > 0) {
    switchToSession(state.sessions[0].id);
  } else {
    createNewSession();
  }
  setStatus(chatStatus, "已删除会话");
}

function autoTitleFromFirstMessage(session: Session): void {
  if (session.title) return;
  const firstUserMsg = session.messages.find((m) => m.role === "user");
  if (firstUserMsg) {
    session.title = firstUserMsg.content.slice(0, 30).replace(/\n/g, " ");
    renderSessionSelect();
  }
}

// ─── Knowledge Base ─────────────────────────────────────────────────────────

type KbStatsResponse = {
  ok: boolean;
  chunkCount: number;
  fileCount: number;
  files: string[];
};

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

async function loadKbFileList(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/kb/files`, { method: "GET" });
    if (!res.ok) {
      kbFileList.innerHTML = '<div class="kb-file-item"><span>加载文件列表失败</span></div>';
      return;
    }

    const data = (await res.json()) as { ok: boolean; files: KbFileEntry[] };
    kbFileList.innerHTML = "";

    if (data.files.length === 0) {
      kbFileList.innerHTML = '<div class="kb-file-item"><span>暂无文件</span></div>';
      return;
    }

    for (const file of data.files) {
      const item = document.createElement("div");
      item.className = "kb-file-item";

      const nameSpan = document.createElement("span");
      nameSpan.className = "kb-file-name";
      nameSpan.textContent = file.fileName;
      nameSpan.title = file.fileName;

      const countSpan = document.createElement("span");
      countSpan.className = "kb-file-count";
      countSpan.textContent = `${file.chunkCount} 块`;

      const deleteBtn = document.createElement("button");
      deleteBtn.textContent = "删除";
      deleteBtn.addEventListener("click", () => {
        void deleteKbFile(file.fileName);
      });

      item.appendChild(nameSpan);
      item.appendChild(countSpan);
      item.appendChild(deleteBtn);
      kbFileList.appendChild(item);
    }
  } catch {
    kbFileList.innerHTML = '<div class="kb-file-item"><span>无法连接本地 Agent</span></div>';
  }
}

async function deleteKbFile(fileName: string): Promise<void> {
  try {
    setStatus(kbStatus, `正在删除 ${fileName}...`);
    const res = await fetch(`${agentBase}/v1/kb/files/${encodeURIComponent(fileName)}`, {
      method: "DELETE",
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(kbStatus, `删除失败: ${err}`);
      return;
    }

    setStatus(kbStatus, `已删除 ${fileName}`);
    await loadKbStats();
    await loadKbFileList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(kbStatus, `删除失败: 无法连接本地 Agent (${message})`);
  }
}

async function clearKb(): Promise<void> {
  if (!confirm("确定要清空整个知识库吗？此操作不可撤销。")) {
    return;
  }

  try {
    setStatus(kbStatus, "正在清空知识库...");
    const res = await fetch(`${agentBase}/v1/kb/clear`, { method: "DELETE" });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(kbStatus, `清空失败: ${err}`);
      return;
    }

    setStatus(kbStatus, "知识库已清空");
    await loadKbStats();
    await loadKbFileList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(kbStatus, `清空失败: 无法连接本地 Agent (${message})`);
  }
}

async function exportKb(): Promise<void> {
  try {
    setStatus(kbStatus, "正在导出...");
    const res = await fetch(`${agentBase}/v1/kb/export`, { method: "GET" });
    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(kbStatus, `导出失败: ${err}`);
      return;
    }

    const data = (await res.json()) as { ok: boolean; chunks: Array<{ id: string; fileName: string; text: string }>; exportedAt: string };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `kb-archive-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    setStatus(kbStatus, `已导出 ${data.chunks.length} 个分块`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(kbStatus, `导出失败: 无法连接本地 Agent (${message})`);
  }
}

async function importKb(): Promise<void> {
  const input = $<HTMLInputElement>("importKbFile");
  if (!input.files || input.files.length === 0) {
    setStatus(kbStatus, "请先选择存档文件");
    return;
  }

  const mode = $<HTMLSelectElement>("importMode").value as "merge" | "replace";
  const importStatus = $<HTMLParagraphElement>("kbImportStatus");

  try {
    setStatus(importStatus, "正在导入...");
    const file = input.files[0];
    const text = await file.text();

    let data: { chunks?: Array<{ id: string; fileName: string; text: string }> };
    try {
      data = JSON.parse(text);
    } catch {
      setStatus(importStatus, "导入失败: 文件不是有效的 JSON");
      return;
    }

    if (!Array.isArray(data.chunks)) {
      setStatus(importStatus, "导入失败: 文件格式不正确，缺少 chunks 字段");
      return;
    }

    const res = await fetch(`${agentBase}/v1/kb/import`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ chunks: data.chunks, mode }),
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(importStatus, `导入失败: ${err}`);
      return;
    }

    const result = (await res.json()) as { ok: boolean; importedChunks: number; totalChunks: number; mode: string };
    const modeLabel = mode === "replace" ? "替换" : "合并";
    setStatus(importStatus, `${modeLabel}导入成功：${result.importedChunks} 个分块，总计 ${result.totalChunks} 个分块`);
    input.value = "";
    await loadKbStats();
    await loadKbFileList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(importStatus, `导入失败: 无法连接本地 Agent (${message})`);
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
    input.value = "";
    await loadKbStats();
    await loadKbFileList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(kbStatus, `上传失败: 无法连接本地 Agent (${message})`);
  }
}

// ─── Provider Config ────────────────────────────────────────────────────────

async function loadProviderConfig(): Promise<void> {
  try {
    const res = await fetch(`${agentBase}/v1/provider/config`, {
      method: "GET",
    });

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
    $<HTMLInputElement>("model").value = data.model || "";
    $<HTMLInputElement>("temperature").value = String(data.temperature ?? 0.2);
    $<HTMLInputElement>("maxTokens").value = String(data.maxTokens ?? 900);
    $<HTMLInputElement>("firstTokenTimeout").value = String(data.firstTokenTimeout ?? 20);
    $<HTMLInputElement>("overallTimeout").value = String(data.overallTimeout ?? 240);

    const apiKeyInput = $<HTMLInputElement>("apiKey");
    apiKeyInput.value = "";
    apiKeyInput.placeholder = data.hasApiKey ? "已保存（留空表示不修改）" : "sk-...";

    setStatus(configStatus, data.hasApiKey ? "已加载已保存配置" : "已加载配置（尚未保存 API Key）");

    // 加载配置后自动刷新模型列表
    void refreshModelList();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `读取配置失败: 无法连接本地 Agent (${message})`);
  }
}

async function refreshModelList(): Promise<void> {
  const modelInput = $<HTMLInputElement>("model");
  const datalist = $<HTMLDataListElement>("modelList");
  const currentModel = modelInput.value.trim();

  setStatus(configStatus, "正在获取模型列表...");

  try {
    const res = await fetch(`${agentBase}/v1/provider/models`, { method: "GET" });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(configStatus, `获取模型列表失败: ${err}`);
      return;
    }

    const data = (await res.json()) as { ok: boolean; models: string[]; currentModel?: string };
    const models = data.models ?? [];

    // 清空并填充 datalist
    datalist.innerHTML = "";
    for (const id of models) {
      const option = document.createElement("option");
      option.value = id;
      datalist.appendChild(option);
    }

    // 恢复当前选中的模型值
    if (currentModel) {
      modelInput.value = currentModel;
    } else if (data.currentModel) {
      modelInput.value = data.currentModel;
    }

    setStatus(configStatus, `已获取 ${models.length} 个模型${models.length > 0 ? "，可从列表选择或手动输入" : ""}`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `获取模型列表失败: 无法连接本地 Agent (${message})`);
  }
}

async function saveProviderConfig(): Promise<void> {
  try {
    setStatus(configStatus, "保存中...");

    const body = {
      baseUrl: $<HTMLInputElement>("baseUrl").value.trim(),
      apiKey: $<HTMLInputElement>("apiKey").value.trim(),
      model: $<HTMLInputElement>("model").value.trim(),
      temperature: Number($<HTMLInputElement>("temperature").value || "0.2"),
      maxTokens: Number($<HTMLInputElement>("maxTokens").value || "900"),
      firstTokenTimeout: Number($<HTMLInputElement>("firstTokenTimeout").value || "20"),
      overallTimeout: Number($<HTMLInputElement>("overallTimeout").value || "240"),
    };

    if (!body.apiKey) {
      setStatus(configStatus, "请填写 API Key（当前版本保存配置时为必填）");
      return;
    }

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
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `保存失败: 无法连接本地 Agent (${message})`);
  }
}

// ─── Word Context (Structured) ───────────────────────────────────────────────

async function getStructuredContext(): Promise<DocumentStructure> {
  const emptyStructure: DocumentStructure = {
    title: "",
    totalParagraphs: 0,
    totalCharacters: 0,
    paragraphs: [],
    selection: { text: "" },
  };

  try {
    if (typeof Word === "undefined") {
      return emptyStructure;
    }

    return Word.run(async (context) => {
      const body = context.document.body;
      const selection = context.document.getSelection();

      // Load paragraphs
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const paraList: DocumentParagraph[] = [];
      const MAX_PARAGRAPHS = 80;
      const MAX_TEXT_LENGTH = 200;

      for (let i = 0; i < Math.min(paragraphs.items.length, MAX_PARAGRAPHS); i++) {
        const p = paragraphs.items[i];
        p.load(["text", "style", "isListItem"]);
      }
      await context.sync();

      for (let i = 0; i < Math.min(paragraphs.items.length, MAX_PARAGRAPHS); i++) {
        const p = paragraphs.items[i];
        const text = (p.text || "").trim();
        if (!text) continue;

        const style = (p.style || "Normal").toString();
        const headingMatch = style.match(/Heading\s*(\d)/i);
        paraList.push({
          index: i,
          text: text.slice(0, MAX_TEXT_LENGTH),
          style,
          headingLevel: headingMatch ? parseInt(headingMatch[1]) : undefined,
          isTable: false,
          isList: p.isListItem,
        });
      }

      // Load selection
      selection.load("text");
      await context.sync();

      // Get total text length (approximate)
      body.load("text");
      await context.sync();

      return {
        title: "",
        totalParagraphs: paragraphs.items.length,
        totalCharacters: (body.text || "").length,
        paragraphs: paraList,
        selection: {
          text: (selection.text || "").slice(0, 2000),
        },
      };
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    console.warn("getStructuredContext failed:", message);
    setStatus(chatStatus, `读取文档结构失败: ${message}`);
    return emptyStructure;
  }
}

// Legacy function for backward compatibility
async function getWordContext(): Promise<{ documentContext: string; selection: string }> {
  try {
    if (typeof Word === "undefined") {
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
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    console.warn("getWordContext failed:", message);
    setStatus(chatStatus, `读取文档上下文失败: ${message}`);
    return { documentContext: "", selection: "" };
  }
}

// ─── Word Action Execution ───────────────────────────────────────────────────

async function applyFormat(
  paragraph: Word.Paragraph,
  format: string,
  context: Word.RequestContext
): Promise<void> {
  if (format === "normal" || !format) return;

  // Word JS API style names — try English first, then Chinese fallback
  // Chinese Word uses localized style names like "标题 1"
  const styleMap: Record<string, string[]> = {
    heading1: ["Heading 1", "标题 1", "标题1"],
    heading2: ["Heading 2", "标题 2", "标题2"],
    heading3: ["Heading 3", "标题 3", "标题3"],
    bullet_list: ["List Paragraph", "列表段落"],
    numbered_list: ["List Paragraph", "列表段落"],
  };

  const candidates = styleMap[format];
  if (!candidates) return;

  for (const styleName of candidates) {
    try {
      paragraph.style = styleName;
      await context.sync();
      return; // Success — exit
    } catch {
      // This style name didn't work, try the next one
      continue;
    }
  }

  console.warn(`无法设置样式 "${format}"，所有候选名称均失败，将使用默认格式`);
}

async function executeAction(action: WordAction): Promise<string> {
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  return Word.run(async (context) => {
    const body = context.document.body;

    switch (action.action) {
      case "insert_after_heading": {
        const { heading_text, content, format = "normal" } = action.params;
        const matches = body.search(heading_text, { matchCase: false, matchWholeWord: false });
        matches.load("items");
        await context.sync();

        if (matches.items.length > 0) {
          const range = matches.items[0];
          const paragraph = range.paragraphs.getFirst();
          paragraph.load("text,style,isListItem");
          const newPara = paragraph.insertParagraph(content, "After");
          await context.sync();
          await applyFormat(newPara, format, context);
          return `已在"${heading_text}"后插入内容`;
        }
        // Fallback: insert at end
        const lastPara = body.paragraphs.getLast();
        const newPara = lastPara.insertParagraph(content, "After");
        await context.sync();
        await applyFormat(newPara, format, context);
        return `未找到标题"${heading_text}"，已插入到文档末尾`;
      }

      case "replace_selection": {
        const { content, format = "normal" } = action.params;
        const selection = context.document.getSelection();
        const para = selection.insertParagraph(content, "Replace");
        await context.sync();
        await applyFormat(para, format, context);
        return "已替换选区内容";
      }

      case "insert_at_end": {
        const { content, format = "normal" } = action.params;
        const lastPara = body.paragraphs.getLast();
        const newPara = lastPara.insertParagraph(content, "After");
        await context.sync();
        await applyFormat(newPara, format, context);
        return "已追加到文档末尾";
      }

      case "insert_at_start": {
        const { content, format = "normal" } = action.params;
        console.log("[DEBUG insert_at_start] content length:", content?.length, "format:", format, "content preview:", JSON.stringify(content?.slice(0, 50)));

        if (!content || String(content).trim().length === 0) {
          throw new Error("插入内容为空或仅包含空白字符 (content is empty)");
        }

        let newPara: Word.Paragraph;
        try {
          const firstPara = body.paragraphs.getFirst();
          firstPara.load("text");
          await context.sync();
          console.log("[DEBUG insert_at_start] firstPara.text:", JSON.stringify(firstPara.text));
          newPara = firstPara.insertParagraph(String(content), "Before");
        } catch (e) {
          const errMsg = e instanceof Error ? e.message : String(e);
          console.log("[DEBUG insert_at_start] getFirst failed, fallback to body.insertParagraph(Start):", errMsg);
          newPara = body.insertParagraph(String(content), "Start");
        }
        await context.sync();
        await applyFormat(newPara, format, context);
        console.log("[DEBUG insert_at_start] success");
        return "已插入到文档开头";
      }

      case "insert_after_paragraph": {
        const { paragraph_index, content, format = "normal" } = action.params;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        if (paragraph_index >= 0 && paragraph_index < paragraphs.items.length) {
          const newPara = paragraphs.items[paragraph_index].insertParagraph(content, "After");
          await context.sync();
          await applyFormat(newPara, format, context);
          return `已在段落 ${paragraph_index} 后插入内容`;
        }
        // Fallback: insert at end
        const lastPara = body.paragraphs.getLast();
        const newPara = lastPara.insertParagraph(content, "After");
        await context.sync();
        await applyFormat(newPara, format, context);
        return `段落序号 ${paragraph_index} 超出范围，已插入到末尾`;
      }

      case "delete_paragraph": {
        const { paragraph_index } = action.params;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        if (paragraph_index >= 0 && paragraph_index < paragraphs.items.length) {
          paragraphs.items[paragraph_index].delete();
          await context.sync();
          return `已删除段落 ${paragraph_index}`;
        }
        return `段落序号 ${paragraph_index} 超出范围，未执行删除`;
      }

      case "find_and_replace": {
        const { find_text, replace_text } = action.params;
        const ranges = body.search(find_text, { matchCase: true, matchWholeWord: false });
        ranges.load("items");
        await context.sync();

        if (ranges.items.length > 0) {
          // Replace first occurrence
          ranges.items[0].insertText(replace_text, "Replace");
          await context.sync();
          return `已替换 ${ranges.items.length} 处匹配中的第 1 处`;
        }
        return `未找到"${find_text}"`;
      }

      case "reply_only":
        // No document operation needed
        return "仅回复文本，无需文档操作";

      default:
        return `未知操作: ${action.action}`;
    }
  });
}

function stringifyOfficeError(error: unknown): { message: string; details: string } {
  let message = "未知错误";
  let details = "";
  if (error instanceof Error) {
    message = error.message;
    details = `name: ${error.name}, message: ${error.message}`;
    // Office.js errors may have extra fields
    const ext = error as any;
    if (ext.code) {
      details += `, code: ${String(ext.code)}`;
    }
    if (ext.traceMessages && Array.isArray(ext.traceMessages)) {
      details += `, traceMessages: ${JSON.stringify(ext.traceMessages)}`;
    }
    if (ext.innerError) {
      details += `, innerError: ${JSON.stringify(ext.innerError)}`;
    }
    if (ext.debugInfo) {
      details += `, debugInfo: ${JSON.stringify(ext.debugInfo)}`;
    }
  } else {
    details = String(error);
  }
  console.error("[OfficeError]", details, error);
  return { message, details };
}

async function executeActionPlan(plan: ActionPlan): Promise<void> {
  const results: string[] = [];

  for (let i = 0; i < plan.actions.length; i++) {
    const action = plan.actions[i];
    try {
      const result = await executeAction(action);
      results.push(`✅ ${action.description}: ${result}`);
    } catch (error) {
      const { message, details } = stringifyOfficeError(error);
      results.push(`❌ ${action.description}: ${message}`);
      console.error(`[executeActionPlan] action ${i + 1} (${action.action}) failed:`, details);

      if (plan.actions.length > 1) {
        const shouldContinue = confirm(`操作 ${i + 1} 失败: ${message}\n\n是否继续执行后续操作？`);
        if (!shouldContinue) break;
      }
    }
  }

  setStatus(chatStatus, results.join("\n"));
}

// ─── Action Plan Preview ────────────────────────────────────────────────────

function showActionPlanPreview(plan: ActionPlan): void {
  state.pendingActionPlan = plan;

  // Build preview content
  let html = `<div class="action-plan-explanation">${escapeHtml(plan.explanation)}</div>`;

  for (const action of plan.actions) {
    html += `<div class="action-plan-item">`;
    html += `<div class="action-plan-action">${escapeHtml(action.description)}</div>`;
    html += `<div class="action-plan-params">`;

    for (const [key, value] of Object.entries(action.params)) {
      if (key === "content") {
        const truncated = String(value).length > 100 ? String(value).slice(0, 100) + "..." : String(value);
        html += `<span class="action-param">${escapeHtml(key)}: "${escapeHtml(truncated)}"</span>`;
      } else {
        html += `<span class="action-param">${escapeHtml(key)}: ${escapeHtml(String(value))}</span>`;
      }
    }

    html += `</div></div>`;
  }

  actionPlanContent.innerHTML = html;
  actionPlanPanel.style.display = "";
  setStatus(chatStatus, `预览：${plan.actions.length} 个操作 — 请确认或取消`);
}

function hideActionPlanPreview(): void {
  actionPlanPanel.style.display = "none";
  state.pendingActionPlan = null;
}

async function confirmActionPlan(): Promise<void> {
  if (!state.pendingActionPlan) {
    setStatus(chatStatus, "没有待执行的操作计划");
    return;
  }

  const plan = state.pendingActionPlan;
  hideActionPlanPreview();

  try {
    await executeActionPlan(plan);
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(chatStatus, `执行操作失败: ${message}`);
  }
}

function escapeHtml(text: string): string {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

// ─── Legacy Insert Functions (kept for backward compatibility) ──────────────

async function applyTextToWordSelection(text: string): Promise<void> {
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(text, "Replace");
    await context.sync();
  });
}

async function insertTextAtCursor(text: string): Promise<void> {
  if (typeof Word === "undefined") {
    throw new Error("当前不在 Word 宿主中");
  }

  const normalized = cleanupMarkdownForWord(text);
  if (!normalized) {
    throw new Error("没有可插入的正文内容");
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(normalized, "End");
    await context.sync();
  });
}

// ─── Preview (legacy, for simple text insert) ────────────────────────────────

const previewPanel = $<HTMLDivElement>("previewPanel");
const previewContent = $<HTMLDivElement>("previewContent");

function showSimplePreview(text: string, mode: "replace" | "append_end" | "insert_start"): void {
  const modeLabels: Record<string, string> = {
    replace: "替换当前选区",
    append_end: "追加到文末",
    insert_start: "插入到文档开头",
  };

  previewContent.textContent = text;
  state.pendingSimpleInsert = { text, mode };
  previewPanel.style.display = "";
  setStatus(chatStatus, `预览：${modeLabels[mode] ?? "插入"} — 请确认或取消`);
}

function hideSimplePreview(): void {
  previewPanel.style.display = "none";
  state.pendingSimpleInsert = null;
}

async function confirmSimplePreview(): Promise<void> {
  if (!state.pendingSimpleInsert) {
    setStatus(chatStatus, "没有待插入的内容");
    return;
  }

  const { text, mode } = state.pendingSimpleInsert;
  hideSimplePreview();

  try {
    if (mode === "replace") {
      await applyTextToWordSelection(text);
      setStatus(chatStatus, "已替换当前选区（已去除 Markdown）");
    } else if (mode === "append_end") {
      await Word.run(async (context) => {
        const lastPara = context.document.body.paragraphs.getLast();
        lastPara.insertParagraph(text, "After");
        await context.sync();
      });
      setStatus(chatStatus, "已追加到文档末尾");
    } else if (mode === "insert_start") {
      await Word.run(async (context) => {
        const body = context.document.body;
        console.log("[DEBUG insert_start simple] text length:", text?.length, "text preview:", JSON.stringify(text?.slice(0, 50)));

        if (!text || text.trim().length === 0) {
          throw new Error("插入内容为空或仅包含空白字符");
        }

        let newPara: Word.Paragraph;
        try {
          const firstPara = body.paragraphs.getFirst();
          firstPara.load("text");
          await context.sync();
          console.log("[DEBUG insert_start simple] firstPara.text:", JSON.stringify(firstPara.text));
          newPara = firstPara.insertParagraph(text, "Before");
        } catch (e) {
          const errMsg = e instanceof Error ? e.message : String(e);
          console.log("[DEBUG insert_start simple] getFirst failed, fallback to body.insertParagraph(Start):", errMsg);
          newPara = body.insertParagraph(text, "Start");
        }
        await context.sync();
        console.log("[DEBUG insert_start simple] inserted successfully");
      });
      setStatus(chatStatus, "已插入到文档开头");
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    console.error("[ERROR confirmSimplePreview] mode:", mode, "error:", message, error);
    setStatus(chatStatus, `插入失败: ${message}`);
  }
}

// ─── Chat Stream ────────────────────────────────────────────────────────────

type StreamEvent = {
  type: string;
  delta?: string;
  reply?: string;
  actionPlan?: ActionPlan | null;
  retrievalCount?: number;
  citations?: Array<{ fileName: string }>;
  error?: string;
};

async function sendMessageStream(
  payload: {
    messages: ChatMessage[];
    documentContext: string;
    documentStructure?: DocumentStructure;
    selection: string;
    insertMode: InsertMode;
  },
  onDelta: (text: string) => void
): Promise<{
  reply: string;
  actionPlan: ActionPlan | null;
  retrievalCount?: number;
  citations?: Array<{ fileName: string }>;
}> {
  const res = await fetch(`${agentBase}/v1/chat/stream`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
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
  let donePayload: { retrievalCount?: number; citations?: Array<{ fileName: string }>; actionPlan?: ActionPlan | null } = {};

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
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

      let data: StreamEvent;
      try {
        data = JSON.parse(raw) as StreamEvent;
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
        fullReply = data.reply ?? fullReply;
        donePayload = {
          retrievalCount: data.retrievalCount,
          citations: data.citations,
          actionPlan: data.actionPlan,
        };
      }
    }
  }

  return { reply: fullReply.trim(), actionPlan: donePayload.actionPlan ?? null, ...donePayload };
}

// ─── Send Message ──────────────────────────────────────────────────────────

async function sendMessage(): Promise<void> {
  const text = userInput.value.trim();
  if (!text) {
    setStatus(chatStatus, "请输入消息");
    return;
  }

  const insertMode = insertModeSelect.value as InsertMode;

  // Ensure we have an active session
  if (!state.activeSessionId) {
    createNewSession();
  }

  const session = getActiveSession();
  if (!session) {
    createNewSession();
  }

  userInput.value = "";
  appendMessage("user", text);

  // Build the actual prompt sent to LLM
  let llmPrompt = text;
  if (insertMode !== "chat_only") {
    llmPrompt = `${text}\n\n请根据文档结构和用户意图选择合适的工具来操作文档。`;
  }

  // Store original user message in session
  const activeSession = getActiveSession()!;
  activeSession.messages.push({ role: "user", content: text });
  autoTitleFromFirstMessage(activeSession);

  setStatus(chatStatus, "思考中（流式返回）...");

  // Get structured document context
  const wordStructure = await getStructuredContext();
  const wordContext = await getWordContext();

  // Show selection info if available
  if (wordStructure.selection.text) {
    appendMessage("system", `已读取选区（${wordStructure.selection.text.length} 字）`);
  }

  try {
    const assistantEl = createAssistantMessage("");
    const data = await sendMessageStream(
      {
        messages: [{ role: "user", content: llmPrompt }],
        documentContext: wordContext.documentContext,
        documentStructure: wordStructure,
        selection: wordContext.selection,
        insertMode,
      },
      (partial) => {
        updateAssistantMessage(assistantEl, partial);
      }
    );

    state.lastReply = data.reply;
    activeSession.messages.push({ role: "assistant", content: data.reply });
    activeSession.updatedAt = Date.now();

    const citationText = data.citations?.length
      ? `\n\n来源: ${data.citations.map((c) => c.fileName).join(", ")}（命中 ${data.retrievalCount ?? data.citations.length} 段）`
      : "";

    // Handle action plan from LLM
    if (data.actionPlan && data.actionPlan.actions.length > 0) {
      updateAssistantMessage(assistantEl, `${data.reply}${citationText}\n\n📋 操作计划：${data.actionPlan.explanation}`);
      showActionPlanPreview(data.actionPlan);
    } else {
      updateAssistantMessage(assistantEl, `${data.reply}${citationText}`);

      // Fallback: if no action plan, handle simple insert modes
      if (insertMode === "replace_selection") {
        const cleanText = cleanupMarkdownForWord(data.reply);
        showSimplePreview(cleanText, "replace");
      } else if (insertMode === "append_end") {
        const cleanText = cleanupMarkdownForWord(data.reply);
        showSimplePreview(cleanText, "append_end");
      } else {
        setStatus(chatStatus, "已完成");
      }
    }

    saveSessionsToStorage();
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(chatStatus, `请求失败: 无法连接本地 Agent (${message})`);
  }
}

// ─── Retry Last Message ─────────────────────────────────────────────────────

async function retryLastMessage(): Promise<void> {
  const session = getActiveSession();
  if (!session || session.messages.length === 0) {
    setStatus(chatStatus, "没有可重试的消息");
    return;
  }

  // Find the last user message
  const lastUserIdx = session.messages.map((m, i) => m.role === "user" ? i : -1).filter((i) => i >= 0).pop();
  if (lastUserIdx === undefined || lastUserIdx < 0) {
    setStatus(chatStatus, "没有可重试的用户消息");
    return;
  }

  const lastUserContent = session.messages[lastUserIdx].content;

  // Remove the last assistant message if it exists after the last user message
  if (lastUserIdx + 1 < session.messages.length && session.messages[lastUserIdx + 1].role === "assistant") {
    session.messages.splice(lastUserIdx + 1);
  }

  // Set the input and send
  userInput.value = lastUserContent;
  await sendMessage();
}

// ─── Insert Last Reply ──────────────────────────────────────────────────────

async function insertLastReplyToSelection(): Promise<void> {
  if (!state.lastReply) {
    setStatus(chatStatus, "没有可插入的回复");
    return;
  }

  const cleanText = cleanupMarkdownForWord(state.lastReply);
  showSimplePreview(cleanText, "replace");
}

async function insertLastReplyAtCursor(): Promise<void> {
  if (!state.lastReply) {
    setStatus(chatStatus, "没有可插入的回复");
    return;
  }

  const cleanText = cleanupMarkdownForWord(state.lastReply);
  showSimplePreview(cleanText, "insert_start");
}

// ─── Bind Actions ────────────────────────────────────────────────────────────

function bindActions(): void {
  // Config
  $("saveConfig").addEventListener("click", () => {
    void saveProviderConfig();
  });

  $("refreshModels").addEventListener("click", () => {
    void refreshModelList();
  });

  // Knowledge base
  $("uploadFile").addEventListener("click", () => {
    void uploadKnowledgeFile();
  });

  $("refreshKb").addEventListener("click", () => {
    void loadKbStats();
    void loadKbFileList();
  });

  $("clearKb").addEventListener("click", () => {
    void clearKb();
  });

  $("exportKb").addEventListener("click", () => {
    void exportKb();
  });

  $("importKbFile").addEventListener("change", () => {
    void importKb();
  });

  // Menu toggle
  $("menuToggle").addEventListener("click", () => {
    const card = document.getElementById("menuCard") as HTMLDetailsElement | null;
    if (card) {
      card.open = !card.open;
    }
  });

  // Chat
  $("sendMsg").addEventListener("click", () => {
    void sendMessage();
  });

  // Allow Enter to send (Shift+Enter for newline)
  userInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      void sendMessage();
    }
  });

  // Insert actions
  $("insertReply").addEventListener("click", () => {
    void insertLastReplyToSelection();
  });

  $("retryLast").addEventListener("click", () => {
    void retryLastMessage();
  });

  $("insertReplyCursor").addEventListener("click", () => {
    void insertLastReplyAtCursor();
  });

  // Simple preview (legacy)
  $("previewConfirm").addEventListener("click", () => {
    void confirmSimplePreview();
  });

  $("previewCancel").addEventListener("click", () => {
    hideSimplePreview();
    setStatus(chatStatus, "已取消插入");
  });

  // Action plan preview
  $("actionPlanConfirm").addEventListener("click", () => {
    void confirmActionPlan();
  });

  $("actionPlanCancel").addEventListener("click", () => {
    hideActionPlanPreview();
    setStatus(chatStatus, "已取消操作计划");
  });

  // Session management
  $("newSession").addEventListener("click", () => {
    createNewSession();
  });

  $("clearSession").addEventListener("click", () => {
    clearCurrentSession();
  });

  $("deleteSession").addEventListener("click", () => {
    deleteCurrentSession();
  });

  sessionSelect.addEventListener("change", () => {
    const selectedId = sessionSelect.value;
    if (selectedId && selectedId !== state.activeSessionId) {
      switchToSession(selectedId);
    }
  });
}

// ─── Initialize ──────────────────────────────────────────────────────────────

Office.onReady(() => {
  loadSessionsFromStorage();

  // Ensure at least one session exists
  if (state.sessions.length === 0) {
    createNewSession();
  } else {
    // Switch to the first session
    switchToSession(state.sessions[0].id);
  }

  bindActions();
  void loadProviderConfig();
  void loadKbStats();
  void loadKbFileList();
  setStatus(chatStatus, "就绪：可直接开始对话");
});
