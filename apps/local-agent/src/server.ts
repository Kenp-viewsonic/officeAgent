import cors from "cors";
import express from "express";
import multer from "multer";
import { z } from "zod";
import { callOpenAICompatible, LlmHttpError } from "./llm.js";
import { keywordRetrieve, splitToChunks } from "./retrieval.js";
import { appendChunks, loadChunks, loadProviderConfig, saveProviderConfig } from "./store.js";
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

const configSchema = z.object({
  baseUrl: z.string().url(),
  apiKey: z.string().min(1),
  model: z.string().min(1),
  temperature: z.number().min(0).max(2).optional(),
  maxTokens: z.number().int().positive().optional(),
});

app.post("/v1/provider/config", async (req, res) => {
  const result = configSchema.safeParse(req.body);
  if (!result.success) {
    return res.status(400).json({ error: "Invalid provider config", details: result.error.flatten() });
  }

  await saveProviderConfig(result.data);
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

app.post("/v1/kb/upload", upload.single("file"), async (req, res) => {
  const file = req.file;
  if (!file) {
    return res.status(400).json({ error: "No file uploaded" });
  }

  const text = file.buffer.toString("utf-8");
  const chunks = splitToChunks(file.originalname, text);
  await appendChunks(chunks);

  return res.json({ ok: true, fileName: file.originalname, chunkCount: chunks.length });
});

const chatSchema = z.object({
  messages: z.array(z.object({ role: z.enum(["system", "user", "assistant"]), content: z.string().min(1) })),
  documentContext: z.string().optional(),
  selection: z.string().optional(),
});

app.post("/v1/chat", async (req, res) => {
  const parsed = chatSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: "Invalid chat payload", details: parsed.error.flatten() });
  }

  const provider = await loadProviderConfig();
  if (!provider) {
    return res.status(400).json({ error: "Provider config missing. Please configure base_url/api_key/model first." });
  }

  const messages = parsed.data.messages as ChatMessage[];
  const userQuery = [...messages].reverse().find((msg) => msg.role === "user")?.content ?? "";

  const chunks = await loadChunks();
  const retrieved = keywordRetrieve(chunks, userQuery, 4);

  const contextParts: string[] = [];
  if (parsed.data.documentContext) {
    contextParts.push(`文档上下文:\n${parsed.data.documentContext}`);
  }
  if (parsed.data.selection) {
    contextParts.push(`当前选区:\n${parsed.data.selection}`);
  }

  const enrichedMessages: ChatMessage[] = contextParts.length
    ? [{ role: "system", content: contextParts.join("\n\n") }, ...messages]
    : messages;

  try {
    const reply = await callOpenAICompatible(provider, enrichedMessages, retrieved);
    return res.json({ ok: true, reply, citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })) });
  } catch (error) {
    if (error instanceof LlmHttpError) {
      const safeStatus = error.status >= 400 && error.status < 600 ? error.status : 502;
      return res.status(safeStatus).json({ error: error.details || error.message });
    }

    const message = error instanceof Error ? error.message : "Unknown error";
    return res.status(502).json({ error: message });
  }
});

app.listen(port, host, () => {
  console.log(`Local agent listening on http://${host}:${port}`);
});
