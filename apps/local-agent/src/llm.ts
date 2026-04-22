import { ChatMessage, ProviderConfig, RetrievalChunk } from "./types.js";

type ChatCompletionResponse = {
  choices?: Array<{ message?: { content?: string } }>;
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

export async function callOpenAICompatible(
  config: ProviderConfig,
  messages: ChatMessage[],
  contextChunks: RetrievalChunk[]
): Promise<string> {
  const contextText = contextChunks
    .map((chunk, idx) => `[${idx + 1}] ${chunk.fileName}: ${chunk.text}`)
    .join("\n\n");

  const systemPrompt: ChatMessage = {
    role: "system",
    content:
      "你是一个面向 Word 文档编辑的本地助手。回答必须尽量可执行，且在引用知识库时标注来源编号。",
  };

  const retrievalPrompt: ChatMessage = {
    role: "system",
    content: contextText
      ? `以下是可用知识片段：\n${contextText}\n\n请尽可能基于这些片段回答。`
      : "当前没有检索到知识库片段，你可以基于用户输入给出通用建议。",
  };

  const payload = {
    model: config.model,
    messages: [systemPrompt, retrievalPrompt, ...messages],
    temperature: config.temperature ?? 0.2,
    max_tokens: config.maxTokens ?? 900,
  };

  const endpoint = config.baseUrl.replace(/\/$/, "") + "/chat/completions";

  let response: Response;
  try {
    response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(30_000),
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
  return data.choices?.[0]?.message?.content?.trim() || "模型没有返回可用内容。";
}
