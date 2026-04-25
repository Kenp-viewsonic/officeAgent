import { ChatMessage, ProviderConfig, RetrievalChunk, ToolDefinition, ActionPlan, WordAction } from "./types.js";

type ChatCompletionResponse = {
  choices?: Array<{
    message?: {
      content?: string;
      tool_calls?: Array<{
        id: string;
        type: "function";
        function: {
          name: string;
          arguments: string;
        };
      }>;
    };
  }>;
};

type StreamChunk = {
  choices?: Array<{ delta?: { content?: string; tool_calls?: any }; text?: string; message?: { content?: string } }>;
};

type ModelsResponse = {
  data?: Array<{ id?: string }>;
};

export class LlmHttpError extends Error {
  status: number;

  details: string;

  constructor(status: number, details: string) {
    super(`LLM request failed (${status}): ${details}`);
    this.name = "LlmHttpError";
    this.status = status;
    this.details = details;
  }
}

// --- Word Tool Definitions ---

export const WORD_TOOLS: ToolDefinition[] = [
  {
    type: "function",
    function: {
      name: "insert_after_heading",
      description: "在指定标题后插入新内容。通过 heading_text 定位标题，支持指定格式。",
      parameters: {
        type: "object",
        properties: {
          heading_text: { type: "string", description: "要查找的标题文本" },
          content: { type: "string", description: "要插入的内容" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "插入内容的格式，默认为 normal",
          },
        },
        required: ["heading_text", "content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "replace_selection",
      description: "替换当前选中的文本内容。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "替换后的新内容" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "替换内容的格式，默认为 normal",
          },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_at_end",
      description: "在文档末尾追加内容。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "要追加的内容" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "追加内容的格式，默认为 normal",
          },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_at_start",
      description: "在文档开头插入内容。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "要插入的内容" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "插入内容的格式，默认为 normal",
          },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_after_paragraph",
      description: "在指定段落序号后插入内容。段落序号从文档结构中获取。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "目标段落的序号（从文档结构中获取）" },
          content: { type: "string", description: "要插入的内容" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "插入内容的格式，默认为 normal",
          },
        },
        required: ["paragraph_index", "content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_paragraph",
      description: "删除指定段落序号的内容。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "要删除的段落序号" },
        },
        required: ["paragraph_index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_and_replace",
      description: "在文档中查找文本并替换为新文本。",
      parameters: {
        type: "object",
        properties: {
          find_text: { type: "string", description: "要查找的文本" },
          replace_text: { type: "string", description: "替换后的文本" },
        },
        required: ["find_text", "replace_text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "reply_only",
      description: "仅回复文本，不执行任何文档操作。用于纯问答、解释、建议等场景。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "回复给用户的文本内容" },
        },
        required: ["content"],
      },
    },
  },
];

// --- Build document structure description ---

export function describeDocumentStructure(structure: {
  totalParagraphs: number;
  paragraphs: Array<{ index: number; text: string; style: string; headingLevel?: number; isList: boolean }>;
  selection: { text: string };
}): string {
  const lines: string[] = [];
  lines.push(`文档共 ${structure.totalParagraphs} 段。结构概览：`);

  for (const p of structure.paragraphs) {
    if (p.headingLevel) {
      lines.push(`  ${"#".repeat(p.headingLevel)} [段落${p.index}] ${p.text}`);
    } else if (p.isList) {
      lines.push(`  - [段落${p.index}] ${p.text}`);
    } else {
      lines.push(`  [段落${p.index}] ${p.text.slice(0, 100)}`);
    }
  }

  if (structure.selection.text) {
    lines.push(`\n当前选区：${structure.selection.text.slice(0, 500)}`);
  }

  return lines.join("\n");
}

// --- Build system prompt based on mode and document structure ---

function buildSystemPrompt(
  hasTools: boolean,
  insertMode: string,
  documentStructureDescription?: string
): ChatMessage {
  let content = `你是一个面向 Word 文档编辑的智能助手。你可以通过调用工具来操作文档。

关键规则：
1. 根据文档结构和用户意图选择最合适的工具和参数
2. 如果用户要求在某个位置插入内容，优先使用 insert_after_heading 或 insert_after_paragraph，而不是让用户手动定位
3. 如果用户要求删除内容，使用 delete_paragraph 或 find_and_replace
4. 如果用户只是提问或需要建议，使用 reply_only
5. 当使用 insert/replace 工具时，content 参数应该是可直接写入 Word 的纯文本，不要使用 Markdown 标记
6. format 参数用于指定插入内容的格式，默认为 normal
7. 在引用知识库时标注来源编号`;

  if (hasTools) {
    content += `

你可以调用以下工具来操作文档：
- insert_after_heading: 在指定标题后插入新内容
- replace_selection: 替换当前选中的文本
- insert_at_end: 在文档末尾追加内容
- insert_at_start: 在文档开头插入内容
- insert_after_paragraph: 在指定段落序号后插入内容
- delete_paragraph: 删除指定段落
- find_and_replace: 查找并替换文本
- reply_only: 仅回复文本，不执行文档操作

请根据用户意图选择合适的工具调用。如果用户只是提问，使用 reply_only。`;
  }

  if (insertMode === "chat_only") {
    content += `

当前模式为"仅对话"，请只使用 reply_only 工具回复用户，不要执行任何文档操作。`;
  } else if (insertMode === "replace_selection") {
    content += `

当前模式为"替换选区"，请优先使用 replace_selection 工具。`;
  } else if (insertMode === "append_end") {
    content += `

当前模式为"追加到文末"，请优先使用 insert_at_end 工具。`;
  } else if (insertMode === "smart_action") {
    content += `

当前模式为"智能操作"，请根据用户意图自主选择最合适的工具。`;
  }

  if (documentStructureDescription) {
    content += `

当前文档结构：
${documentStructureDescription}`;
  }

  return { role: "system", content };
}

// --- Build payload ---

function buildPayload(
  config: ProviderConfig,
  messages: ChatMessage[],
  contextChunks: RetrievalChunk[],
  stream: boolean,
  tools?: ToolDefinition[],
  insertMode?: string,
  documentStructureDescription?: string
): Record<string, any> {
  const contextText = contextChunks.map((chunk, idx) => `[${idx + 1}] ${chunk.fileName}: ${chunk.text}`).join("\n\n");

  const hasTools = !!(tools && tools.length > 0);
  const systemPrompt = buildSystemPrompt(hasTools, insertMode || "smart_action", documentStructureDescription);

  const retrievalPrompt: ChatMessage = {
    role: "system",
    content: contextText
      ? `以下是可用知识片段：\n${contextText}\n\n请尽可能基于这些片段回答，并标注来源编号。`
      : "当前没有检索到知识库片段，你可以基于用户输入给出通用建议。",
  };

  const payload: Record<string, any> = {
    model: config.model,
    messages: [systemPrompt, retrievalPrompt, ...messages],
    temperature: config.temperature ?? 0.2,
    max_tokens: config.maxTokens ?? 900,
    stream,
  };

  if (hasTools) {
    payload.tools = tools;
    payload.tool_choice = "auto";
  }

  return payload;
}

// --- Parse action plan from LLM response ---

export function parseActionPlanFromToolCalls(
  toolCalls: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }>
): ActionPlan {
  const actions: WordAction[] = [];

  for (const tc of toolCalls) {
    let params: Record<string, any> = {};
    try {
      params = JSON.parse(tc.function.arguments);
    } catch {
      params = { raw_arguments: tc.function.arguments };
    }

    // Generate human-readable description
    const desc = describeAction(tc.function.name, params);

    actions.push({
      action: tc.function.name,
      params,
      description: desc,
    });
  }

  return {
    actions,
    explanation: actions.map((a) => a.description).join("；"),
  };
}

function describeAction(actionName: string, params: Record<string, any>): string {
  switch (actionName) {
    case "insert_after_heading":
      return `在标题"${params.heading_text}"后插入内容${params.format && params.format !== "normal" ? `（格式: ${params.format}）` : ""}`;
    case "replace_selection":
      return `替换选区内容${params.format && params.format !== "normal" ? `（格式: ${params.format}）` : ""}`;
    case "insert_at_end":
      return `追加内容到文档末尾${params.format && params.format !== "normal" ? `（格式: ${params.format}）` : ""}`;
    case "insert_at_start":
      return `插入内容到文档开头${params.format && params.format !== "normal" ? `（格式: ${params.format}）` : ""}`;
    case "insert_after_paragraph":
      return `在段落${params.paragraph_index}后插入内容${params.format && params.format !== "normal" ? `（格式: ${params.format}）` : ""}`;
    case "delete_paragraph":
      return `删除段落${params.paragraph_index}`;
    case "find_and_replace":
      return `将"${params.find_text}"替换为"${params.replace_text}"`;
    case "reply_only":
      return "仅回复文本";
    default:
      return `${actionName}: ${JSON.stringify(params)}`;
  }
}

// --- Parse action plan from text (fallback for non-FC models) ---

export function parseActionPlanFromText(text: string): ActionPlan | null {
  // Try to extract <action_plan>...</action_plan> block
  const planMatch = text.match(/<action_plan>\s*([\s\S]*?)\s*<\/action_plan>/);
  if (!planMatch) {
    return null;
  }

  try {
    const plan = JSON.parse(planMatch[1].trim());
    if (!plan.actions || !Array.isArray(plan.actions)) {
      return null;
    }

    const actions: WordAction[] = plan.actions.map((a: any) => ({
      action: a.action || a.name || "reply_only",
      params: a.params || a.parameters || {},
      description: a.description || describeAction(a.action || a.name || "reply_only", a.params || a.parameters || {}),
    }));

    return {
      actions,
      explanation: plan.explanation || actions.map((a) => a.description).join("；"),
    };
  } catch {
    return null;
  }
}

// --- Extract pure text reply (strip <action_plan> blocks) ---

export function extractTextReply(text: string): string {
  return text.replace(/<action_plan>[\s\S]*?<\/action_plan>/g, "").trim();
}

// --- API helpers ---

function getEndpoint(config: ProviderConfig): string {
  return config.baseUrl.replace(/\/$/, "") + "/chat/completions";
}

function getModelsEndpoint(config: ProviderConfig): string {
  return config.baseUrl.replace(/\/$/, "") + "/models";
}

export async function listOpenAICompatibleModels(config: ProviderConfig): Promise<string[]> {
  const endpoint = getModelsEndpoint(config);

  let response: Response;
  try {
    response = await fetch(endpoint, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${config.apiKey}`,
      },
      signal: AbortSignal.timeout(20_000),
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown network error";
    throw new Error(`LLM network error to ${endpoint}: ${message}`);
  }

  if (!response.ok) {
    const details = await response.text();
    throw new LlmHttpError(response.status, details);
  }

  const data = (await response.json()) as ModelsResponse;
  return (data.data ?? []).map((item) => item.id).filter((id): id is string => Boolean(id));
}

export async function callOpenAICompatible(
  config: ProviderConfig,
  messages: ChatMessage[],
  contextChunks: RetrievalChunk[],
  tools?: ToolDefinition[],
  insertMode?: string,
  documentStructureDescription?: string
): Promise<{ reply: string; actionPlan: ActionPlan | null }> {
  const payload = buildPayload(config, messages, contextChunks, false, tools, insertMode, documentStructureDescription);
  const endpoint = getEndpoint(config);
  const timeoutMs = Number(process.env.LLM_TIMEOUT_MS ?? 180_000);

  let response: Response;
  try {
    response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "text/event-stream, application/json",
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(timeoutMs),
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown network error";
    const causeCode =
      error && typeof error === "object" && "cause" in error
        ? (error as { cause?: { code?: string } }).cause?.code
        : undefined;
    const codeSuffix = causeCode ? ` (${causeCode})` : "";
    throw new Error(`LLM network error to ${endpoint}: ${message}${codeSuffix}`);
  }

  if (!response.ok) {
    const details = await response.text();
    throw new LlmHttpError(response.status, details);
  }

  const data = (await response.json()) as ChatCompletionResponse;
  const choice = data.choices?.[0];
  const message = choice?.message;

  if (!message) {
    return { reply: "模型没有返回可用内容。", actionPlan: null };
  }

  // Check for tool_calls
  if (message.tool_calls && message.tool_calls.length > 0) {
    const actionPlan = parseActionPlanFromToolCalls(message.tool_calls);
    // If there's also text content, include it as the reply
    const reply = message.content?.trim() || actionPlan.explanation;
    return { reply, actionPlan };
  }

  // No tool_calls — check for text-based action plan (fallback)
  const textContent = message.content?.trim() || "模型没有返回可用内容。";
  const textActionPlan = parseActionPlanFromText(textContent);
  const reply = textActionPlan ? extractTextReply(textContent) || textActionPlan.explanation : textContent;

  return { reply, actionPlan: textActionPlan };
}

export async function streamOpenAICompatible(
  config: ProviderConfig,
  messages: ChatMessage[],
  contextChunks: RetrievalChunk[],
  onDelta: (text: string) => void,
  signal?: AbortSignal,
  tools?: ToolDefinition[],
  insertMode?: string,
  documentStructureDescription?: string
): Promise<{ reply: string; actionPlan: ActionPlan | null }> {
  const payload = buildPayload(config, messages, contextChunks, true, tools, insertMode, documentStructureDescription);
  const endpoint = getEndpoint(config);
  const overallTimeoutMs = config.overallTimeout ? config.overallTimeout * 1000 : Number(process.env.LLM_STREAM_TIMEOUT_MS ?? 240_000);
  const firstTokenTimeoutMs = config.firstTokenTimeout ? config.firstTokenTimeout * 1000 : Number(process.env.LLM_STREAM_FIRST_TOKEN_TIMEOUT_MS ?? 20_000);

  const controller = new AbortController();
  let externalAbortHandler: (() => void) | null = null;
  if (signal) {
    if (signal.aborted) {
      controller.abort(signal.reason);
    } else {
      externalAbortHandler = () => controller.abort(signal.reason);
      signal.addEventListener("abort", externalAbortHandler, { once: true });
    }
  }

  const overallTimer = setTimeout(() => {
    controller.abort(new Error("stream_overall_timeout"));
  }, overallTimeoutMs);

  let firstTokenReceived = false;
  let firstTokenTimer = setTimeout(() => {
    if (!firstTokenReceived) {
      controller.abort(new Error("stream_first_token_timeout"));
    }
  }, firstTokenTimeoutMs);

  let response: Response;
  try {
    response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });
  } catch (error) {
    clearTimeout(firstTokenTimer);
    clearTimeout(overallTimer);
    if (signal && externalAbortHandler) {
      signal.removeEventListener("abort", externalAbortHandler);
    }

    if (error instanceof Error && controller.signal.aborted) {
      const reason = controller.signal.reason;
      if (reason instanceof Error && reason.message === "stream_first_token_timeout") {
        throw new Error(`LLM stream timeout: first token not received within ${firstTokenTimeoutMs}ms`);
      }
      if (reason instanceof Error && reason.message === "stream_overall_timeout") {
        throw new Error(`LLM stream timeout: response not finished within ${overallTimeoutMs}ms`);
      }
    }

    const message = error instanceof Error ? error.message : "Unknown network error";
    const causeCode =
      error && typeof error === "object" && "cause" in error
        ? (error as { cause?: { code?: string } }).cause?.code
        : undefined;
    const codeSuffix = causeCode ? ` (${causeCode})` : "";
    throw new Error(`LLM network error to ${endpoint}: ${message}${codeSuffix}`);
  }

  if (!response.ok) {
    clearTimeout(firstTokenTimer);
    clearTimeout(overallTimer);
    if (signal && externalAbortHandler) {
      signal.removeEventListener("abort", externalAbortHandler);
    }
    const details = await response.text();
    throw new LlmHttpError(response.status, details);
  }

  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    clearTimeout(firstTokenTimer);
    clearTimeout(overallTimer);
    if (signal && externalAbortHandler) {
      signal.removeEventListener("abort", externalAbortHandler);
    }

    const data = (await response.json()) as ChatCompletionResponse;
    const choice = data.choices?.[0];
    const message = choice?.message;

    if (!message) {
      return { reply: "模型没有返回可用内容。", actionPlan: null };
    }

    // Check for tool_calls in non-streaming JSON response
    if (message.tool_calls && message.tool_calls.length > 0) {
      const actionPlan = parseActionPlanFromToolCalls(message.tool_calls);
      const reply = message.content?.trim() || actionPlan.explanation;
      if (reply) {
        onDelta(reply);
      }
      return { reply, actionPlan };
    }

    const textContent = message.content?.trim() || "模型没有返回可用内容。";
    const textActionPlan = parseActionPlanFromText(textContent);
    const reply = textActionPlan ? extractTextReply(textContent) || textActionPlan.explanation : textContent;
    if (textContent) {
      onDelta(textContent);
    }
    return { reply, actionPlan: textActionPlan };
  }

  if (!response.body) {
    clearTimeout(firstTokenTimer);
    clearTimeout(overallTimer);
    if (signal && externalAbortHandler) {
      signal.removeEventListener("abort", externalAbortHandler);
    }
    throw new Error("LLM stream response has no body.");
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let buffer = "";
  let fullText = "";
  let pendingToolCalls: Map<number, { id: string; name: string; arguments: string }> = new Map();

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }

    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split(/\r?\n/);
    buffer = lines.pop() ?? "";

    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line.startsWith("data:")) {
        continue;
      }

      const data = line.slice(5).trim();
      if (!data || data === "[DONE]") {
        continue;
      }

      let parsed: StreamChunk;
      try {
        parsed = JSON.parse(data) as StreamChunk;
      } catch {
        continue;
      }

      const delta = parsed.choices?.[0]?.delta;
      if (!delta) {
        // Try legacy format
        const text = parsed.choices?.[0]?.text ?? parsed.choices?.[0]?.message?.content ?? "";
        if (text) {
          if (!firstTokenReceived) {
            firstTokenReceived = true;
            clearTimeout(firstTokenTimer);
          }
          fullText += text;
          onDelta(text);
        }
        continue;
      }

      // Handle tool_calls in streaming
      if (delta.tool_calls && Array.isArray(delta.tool_calls)) {
        for (const tc of delta.tool_calls as Array<{ index?: number; id?: string; function?: { name?: string; arguments?: string } }>) {
          const idx = tc.index ?? 0;
          const existing = pendingToolCalls.get(idx);
          if (tc.id) {
            // New tool call starting
            pendingToolCalls.set(idx, {
              id: tc.id,
              name: tc.function?.name || (existing?.name ?? ""),
              arguments: tc.function?.arguments || "",
            });
          } else if (existing) {
            // Continuation of tool call
            if (tc.function?.name) {
              existing.name += tc.function.name;
            }
            if (tc.function?.arguments) {
              existing.arguments += tc.function.arguments;
            }
          }
        }
        continue;
      }

      // Handle text content
      const content = delta.content ?? "";
      if (content) {
        if (!firstTokenReceived) {
          firstTokenReceived = true;
          clearTimeout(firstTokenTimer);
        }
        fullText += content;
        onDelta(fullText);
      }
    }
  }

  clearTimeout(firstTokenTimer);
  clearTimeout(overallTimer);
  if (signal && externalAbortHandler) {
    signal.removeEventListener("abort", externalAbortHandler);
  }

  // Process accumulated tool calls
  if (pendingToolCalls.size > 0) {
    const toolCalls = [...pendingToolCalls.entries()]
      .sort(([a], [b]) => a - b)
      .map(([, tc]) => ({
        id: tc.id,
        type: "function" as const,
        function: {
          name: tc.name,
          arguments: tc.arguments,
        },
      }));

    const actionPlan = parseActionPlanFromToolCalls(toolCalls);
    const reply = fullText.trim() || actionPlan.explanation;
    return { reply, actionPlan };
  }

  // No tool calls — check for text-based action plan
  const textActionPlan = parseActionPlanFromText(fullText);
  const reply = textActionPlan ? extractTextReply(fullText) || textActionPlan.explanation : fullText.trim() || "模型没有返回可用内容。";
  return { reply, actionPlan: textActionPlan };
}
