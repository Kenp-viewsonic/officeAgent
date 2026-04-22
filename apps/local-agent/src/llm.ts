import { ChatMessage, ProviderConfig, RetrievalChunk } from "./types.js";

type ChatCompletionResponse = {
  choices?: Array<{ message?: { content?: string } }>;
};

type StreamChunk = {
  choices?: Array<{ delta?: { content?: string }; text?: string; message?: { content?: string } }>;
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

function buildPayload(config: ProviderConfig, messages: ChatMessage[], contextChunks: RetrievalChunk[], stream: boolean) {
  const contextText = contextChunks.map((chunk, idx) => `[${idx + 1}] ${chunk.fileName}: ${chunk.text}`).join("\n\n");

  const systemPrompt: ChatMessage = {
    role: "system",
    content: "你是一个面向 Word 文档编辑的本地助手。回答必须尽量可执行，且在引用知识库时标注来源编号。",
  };

  const retrievalPrompt: ChatMessage = {
    role: "system",
    content: contextText
      ? `以下是可用知识片段：\n${contextText}\n\n请尽可能基于这些片段回答。`
      : "当前没有检索到知识库片段，你可以基于用户输入给出通用建议。",
  };

  return {
    model: config.model,
    messages: [systemPrompt, retrievalPrompt, ...messages],
    temperature: config.temperature ?? 0.2,
    max_tokens: config.maxTokens ?? 900,
    stream,
  };
}

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
  contextChunks: RetrievalChunk[]
): Promise<string> {
  const payload = buildPayload(config, messages, contextChunks, false);
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
  return data.choices?.[0]?.message?.content?.trim() || "模型没有返回可用内容。";
}

export async function streamOpenAICompatible(
  config: ProviderConfig,
  messages: ChatMessage[],
  contextChunks: RetrievalChunk[],
  onDelta: (text: string) => void,
  signal?: AbortSignal
): Promise<string> {
  const payload = buildPayload(config, messages, contextChunks, true);
  const endpoint = getEndpoint(config);
  const overallTimeoutMs = Number(process.env.LLM_STREAM_TIMEOUT_MS ?? 240_000);
  const firstTokenTimeoutMs = Number(process.env.LLM_STREAM_FIRST_TOKEN_TIMEOUT_MS ?? 20_000);

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
    const text = data.choices?.[0]?.message?.content?.trim() || "模型没有返回可用内容。";
    if (text) {
      onDelta(text);
    }
    return text;
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

      const delta =
        parsed.choices?.[0]?.delta?.content ??
        parsed.choices?.[0]?.text ??
        parsed.choices?.[0]?.message?.content ??
        "";
      if (!delta) {
        continue;
      }

      if (!firstTokenReceived) {
        firstTokenReceived = true;
        clearTimeout(firstTokenTimer);
      }

      fullText += delta;
      onDelta(delta);
    }
  }

  clearTimeout(firstTokenTimer);
  clearTimeout(overallTimer);
  if (signal && externalAbortHandler) {
    signal.removeEventListener("abort", externalAbortHandler);
  }

  return fullText.trim() || "模型没有返回可用内容。";
}
