import cors from "cors";
import express from "express";
import multer from "multer";
import path from "node:path";
import mammoth from "mammoth";
import { z } from "zod";
import { callOpenAICompatible, listOpenAICompatibleModels, LlmHttpError, streamOpenAICompatible } from "./llm.js";
import { keywordRetrieve, splitToChunks } from "./retrieval.js";
import { appendChunks, getKbStats, loadChunks, loadProviderConfig, saveProviderConfig } from "./store.js";
import { ChatMessage } from "./types.js";

const app = express();
const upload = multer({ limits: { fileSize: Number(process.env.MAX_UPLOAD_MB ?? 20) * 1024 * 1024 } });

const host = process.env.AGENT_BIND_HOST ?? "127.0.0.1";
const port = Number(process.env.PORT ?? 8787);

app.use(cors({ origin: true }));
app.use(express.json({ limit: "2mb" }));

app.get("/health", (_req, res) => {
  res.json({ ok: true, service: "office-agent-local", host, port });
});

function isUtf8TextExt(fileName: string): boolean {
  const ext = path.extname(fileName).toLowerCase();
  return [
    ".txt",
    ".md",
    ".markdown",
    ".json",
    ".csv",
    ".log",
    ".xml",
    ".html",
    ".htm",
    ".ts",
    ".tsx",
    ".js",
    ".jsx",
    ".py",
    ".java",
    ".c",
    ".cpp",
    ".h",
    ".hpp",
    ".go",
    ".rs",
    ".sql",
    ".yaml",
    ".yml",
  ].includes(ext);
}

async function extractUploadText(file: Express.Multer.File): Promise<string> {
  const lowerName = file.originalname.toLowerCase();
  if (lowerName.endsWith(".docx")) {
    const result = await mammoth.extractRawText({ buffer: file.buffer });
    return result.value;
  }

  if (file.mimetype.startsWith("text/") || isUtf8TextExt(file.originalname)) {
    return file.buffer.toString("utf-8");
  }

  throw new Error(`暂不支持该文件类型: ${file.originalname} (${file.mimetype || "unknown"})`);
}

const configSchema = z.object({
  baseUrl: z.string().url(),
  apiKey: z.string().min(1).optional(),
  model: z.string().min(1),
  temperature: z.number().min(0).max(2).optional(),
  maxTokens: z.number().int().positive().optional(),
});

app.post("/v1/provider/config", async (req, res) => {
  const result = configSchema.safeParse(req.body);
  if (!result.success) {
    return res.status(400).json({ error: "Invalid provider config", details: result.error.flatten() });
  }

  const current = await loadProviderConfig();
  const resolvedApiKey = result.data.apiKey || current?.apiKey;
  if (!resolvedApiKey) {
    return res.status(400).json({ error: "API Key is required for first-time save." });
  }

  await saveProviderConfig({ ...result.data, apiKey: resolvedApiKey });
  return res.json({ ok: true });
});

app.get("/v1/provider/config", async (_req, res) => {
  const config = await loadProviderConfig();
  if (!config) {
    return res.status(404).json({ error: "Provider config not found" });
  }

  return res.json({
    baseUrl: config.baseUrl,
    model: config.model,
    temperature: config.temperature,
    maxTokens: config.maxTokens,
    hasApiKey: Boolean(config.apiKey),
  });
});

app.get("/v1/provider/models", async (_req, res) => {
  const config = await loadProviderConfig();
  if (!config) {
    return res.status(400).json({ error: "Provider config missing. Please save base_url/api_key first." });
  }

  try {
    const models = await listOpenAICompatibleModels(config);
    return res.json({ ok: true, models, currentModel: config.model });
  } catch (error) {
    if (error instanceof LlmHttpError) {
      const safeStatus = error.status >= 400 && error.status < 600 ? error.status : 502;
      return res.status(safeStatus).json({ error: error.details || error.message });
    }

    const message = error instanceof Error ? error.message : "Unknown error";
    return res.status(502).json({ error: message });
  }
});

app.get("/v1/kb/stats", async (_req, res) => {
  const stats = await getKbStats();
  return res.json({ ok: true, ...stats });
});

app.post("/v1/kb/upload", upload.single("file"), async (req, res) => {
  const file = req.file;
  if (!file) {
    return res.status(400).json({ error: "No file uploaded" });
  }

  let text: string;
  try {
    text = await extractUploadText(file);
  } catch (error) {
    const message = error instanceof Error ? error.message : "文件解析失败";
    return res.status(400).json({ error: message });
  }

  const chunks = splitToChunks(file.originalname, text);
  if (chunks.length === 0) {
    return res.status(400).json({ error: "文件未提取到可索引文本，请检查文件内容或编码。" });
  }

  await appendChunks(chunks);

  return res.json({ ok: true, fileName: file.originalname, chunkCount: chunks.length });
});

const chatSchema = z.object({
  messages: z.array(z.object({ role: z.enum(["system", "user", "assistant"]), content: z.string().min(1) })),
  documentContext: z.string().optional(),
  selection: z.string().optional(),
});

async function buildChatContext(payload: z.infer<typeof chatSchema>) {
  const provider = await loadProviderConfig();
  if (!provider) {
    throw new Error("Provider config missing. Please configure base_url/api_key/model first.");
  }

  const messages = payload.messages as ChatMessage[];
  const userQuery = [...messages].reverse().find((msg) => msg.role === "user")?.content ?? "";

  const chunks = await loadChunks();
  const retrieved = keywordRetrieve(chunks, userQuery, 4);

  const contextParts: string[] = [];
  if (payload.documentContext) {
    contextParts.push(`文档上下文:\n${payload.documentContext}`);
  }
  if (payload.selection) {
    contextParts.push(`当前选区:\n${payload.selection}`);
  }

  const enrichedMessages: ChatMessage[] = contextParts.length
    ? [{ role: "system", content: contextParts.join("\n\n") }, ...messages]
    : messages;

  return { provider, enrichedMessages, retrieved };
}

app.post("/v1/chat", async (req, res) => {
  const parsed = chatSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: "Invalid chat payload", details: parsed.error.flatten() });
  }

  try {
    const { provider, enrichedMessages, retrieved } = await buildChatContext(parsed.data);
    const reply = await callOpenAICompatible(provider, enrichedMessages, retrieved);
    return res.json({
      ok: true,
      reply,
      retrievalCount: retrieved.length,
      citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })),
    });
  } catch (error) {
    if (error instanceof Error && error.message.startsWith("Provider config missing")) {
      return res.status(400).json({ error: error.message });
    }

    if (error instanceof LlmHttpError) {
      const safeStatus = error.status >= 400 && error.status < 600 ? error.status : 502;
      return res.status(safeStatus).json({ error: error.details || error.message });
    }

    const message = error instanceof Error ? error.message : "Unknown error";
    return res.status(502).json({ error: message });
  }
});

app.post("/v1/chat/stream", async (req, res) => {
  const parsed = chatSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: "Invalid chat payload", details: parsed.error.flatten() });
  }

  const abortController = new AbortController();
  req.on("aborted", () => {
    abortController.abort(new Error("client_aborted"));
  });
  res.on("close", () => {
    // Close can happen both on normal completion and client disconnect.
    if (!res.writableEnded && !abortController.signal.aborted) {
      abortController.abort(new Error("client_closed"));
    }
  });

  try {
    const { provider, enrichedMessages, retrieved } = await buildChatContext(parsed.data);

    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders();

    // Send an immediate event to confirm stream startup to the client.
    res.write(`data: ${JSON.stringify({ type: "start", ts: Date.now() })}\n\n`);

    let reply: string;
    try {
      reply = await streamOpenAICompatible(
        provider,
        enrichedMessages,
        retrieved,
        (delta) => {
          res.write(`data: ${JSON.stringify({ type: "delta", delta })}\n\n`);
        },
        abortController.signal
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      const shouldFallbackToNonStream = message.includes("first token not received");

      if (!shouldFallbackToNonStream) {
        throw error;
      }

      // Some OpenAI-compatible providers accept stream=true but don't emit SSE chunks.
      // Fall back to non-stream completion so client can still receive a final answer.
      reply = await callOpenAICompatible(provider, enrichedMessages, retrieved);
      res.write(`data: ${JSON.stringify({ type: "fallback", reason: "no_stream_delta" })}\n\n`);
    }

    res.write(`data: ${JSON.stringify({
      type: "done",
      reply,
      retrievalCount: retrieved.length,
      citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })),
    })}\n\n`);
    return res.end();
  } catch (error) {
    if (abortController.signal.aborted) {
      return;
    }

    if (!res.headersSent) {
      if (error instanceof Error && error.message.startsWith("Provider config missing")) {
        return res.status(400).json({ error: error.message });
      }

      if (error instanceof LlmHttpError) {
        const safeStatus = error.status >= 400 && error.status < 600 ? error.status : 502;
        return res.status(safeStatus).json({ error: error.details || error.message });
      }

      const message = error instanceof Error ? error.message : "Unknown error";
      return res.status(502).json({ error: message });
    }

    const message = error instanceof Error ? error.message : "Unknown error";
    res.write(`data: ${JSON.stringify({ type: "error", error: message })}\n\n`);
    return res.end();
  }
});

app.listen(port, host, () => {
  console.log(`Local agent listening on http://${host}:${port}`);
});
