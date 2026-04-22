const agentBase = "/api";

type ChatMessage = {
  role: "system" | "user" | "assistant";
  content: string;
};

const state: {
  messages: ChatMessage[];
  lastReply: string;
} = {
  messages: [],
  lastReply: "",
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
const chatStatus = $<HTMLParagraphElement>("chatStatus");
const chatLog = $<HTMLDivElement>("chatLog");

type ProviderConfigView = {
  baseUrl: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
  hasApiKey: boolean;
};

function setStatus(target: HTMLParagraphElement, text: string): void {
  target.textContent = text;
}

function appendMessage(role: "user" | "assistant", content: string): void {
  const div = document.createElement("div");
  div.className = `msg ${role}`;
  div.textContent = `${role === "user" ? "你" : "助手"}: ${content}`;
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
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

    const apiKeyInput = $<HTMLInputElement>("apiKey");
    apiKeyInput.value = "";
    apiKeyInput.placeholder = data.hasApiKey ? "已保存（留空表示不修改）" : "sk-...";

    setStatus(configStatus, data.hasApiKey ? "已加载已保存配置" : "已加载配置（尚未保存 API Key）");
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(configStatus, `读取配置失败: 无法连接本地 Agent (${message})`);
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

async function sendMessage(): Promise<void> {
  const input = $<HTMLTextAreaElement>("userInput");
  const text = input.value.trim();
  if (!text) {
    setStatus(chatStatus, "请输入消息");
    return;
  }

  input.value = "";
  appendMessage("user", text);
  state.messages.push({ role: "user", content: text });
  setStatus(chatStatus, "思考中...");

  const wordContext = await getWordContext();

  try {
    const res = await fetch(`${agentBase}/v1/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        messages: state.messages,
        documentContext: wordContext.documentContext,
        selection: wordContext.selection,
      }),
    });

    if (!res.ok) {
      const err = await parseErrorMessage(res);
      setStatus(chatStatus, `请求失败: ${err}`);
      return;
    }

    const data = (await res.json()) as { reply: string; citations?: Array<{ fileName: string }> };
    state.lastReply = data.reply;
    state.messages.push({ role: "assistant", content: data.reply });

    const citationText = data.citations?.length
      ? `\n\n来源: ${data.citations.map((c) => c.fileName).join(", ")}`
      : "";

    appendMessage("assistant", `${data.reply}${citationText}`);
    setStatus(chatStatus, "已完成");
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    setStatus(chatStatus, `请求失败: 无法连接本地 Agent (${message})`);
  }
}

async function insertLastReplyToWord(): Promise<void> {
  if (!state.lastReply) {
    setStatus(chatStatus, "没有可插入的回复");
    return;
  }

  if (!(window as any).Word) {
    setStatus(chatStatus, "当前不在 Word 宿主中");
    return;
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(state.lastReply, "Replace");
    await context.sync();
  });

  setStatus(chatStatus, "已插入到当前光标位置");
}

function bindActions(): void {
  $("saveConfig").addEventListener("click", () => {
    void saveProviderConfig();
  });

  $("uploadFile").addEventListener("click", () => {
    void uploadKnowledgeFile();
  });

  $("sendMsg").addEventListener("click", () => {
    void sendMessage();
  });

  $("insertReply").addEventListener("click", () => {
    void insertLastReplyToWord();
  });
}

Office.onReady(() => {
  bindActions();
  void loadProviderConfig();
  setStatus(chatStatus, "就绪：可直接开始对话");
});
