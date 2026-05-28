import cors from "cors";
import express from "express";
import multer from "multer";
import path from "node:path";
import mammoth from "mammoth";
import { z } from "zod";
import { createSession, getSession, appendToSession, deleteSession } from "./agent-session.js";
import { callOpenAICompatible, describeDocumentStructure, isPerceptionOnlyPlan, listOpenAICompatibleModels, LlmHttpError, streamOpenAICompatible, WORD_TOOLS } from "./llm.js";
import { keywordRetrieve, splitToChunks } from "./retrieval.js";
import { appendChunks, clearAllChunks, deleteChunksByFile, getKbFileList, getKbStats, importChunks, loadChunks, loadProviderConfig, saveProviderConfig } from "./store.js";
import { ActionPlan, ChatMessage, DocumentStructure, ProviderConfig, RetrievalChunk } from "./types.js";

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
  firstTokenTimeout: z.number().int().min(5).max(120).optional(),
  overallTimeout: z.number().int().min(30).max(600).optional(),
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
    firstTokenTimeout: config.firstTokenTimeout,
    overallTimeout: config.overallTimeout,
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

app.get("/v1/kb/files", async (_req, res) => {
  const files = await getKbFileList();
  return res.json({ ok: true, files });
});

app.delete("/v1/kb/files/:fileName", async (req, res) => {
  const fileName = decodeURIComponent(req.params.fileName);
  if (!fileName) {
    return res.status(400).json({ error: "fileName is required" });
  }

  const removed = await deleteChunksByFile(fileName);
  if (removed === 0) {
    return res.status(404).json({ error: `File not found: ${fileName}` });
  }

  return res.json({ ok: true, fileName, removedChunks: removed });
});

app.delete("/v1/kb/clear", async (_req, res) => {
  const removed = await clearAllChunks();
  return res.json({ ok: true, removedChunks: removed });
});

app.get("/v1/kb/export", async (_req, res) => {
  const chunks = await loadChunks();
  return res.json({ ok: true, chunks, exportedAt: new Date().toISOString() });
});

const importSchema = z.object({
  chunks: z.array(z.object({
    id: z.string(),
    fileName: z.string(),
    text: z.string(),
  })),
  mode: z.enum(["merge", "replace"]).optional(),
});

app.post("/v1/kb/import", async (req, res) => {
  const result = importSchema.safeParse(req.body);
  if (!result.success) {
    return res.status(400).json({ error: "Invalid import payload", details: result.error.flatten() });
  }

  const { chunks, mode = "merge" } = result.data;

  try {
    const { importedChunks, totalChunks } = await importChunks(chunks, mode as "merge" | "replace");
    return res.json({ ok: true, importedChunks, totalChunks, mode });
  } catch (error) {
    const message = error instanceof Error ? error.message : "导入失败";
    return res.status(500).json({ error: message });
  }
});

// --- Chat Schema with document structure and insert mode ---

const documentStructureSchema = z.object({
  title: z.string(),
  totalParagraphs: z.number(),
  totalCharacters: z.number(),
  paragraphs: z.array(z.object({
    index: z.number(),
    text: z.string(),
    style: z.string(),
    headingLevel: z.number().optional(),
    isTable: z.boolean(),
    isList: z.boolean(),
    charCount: z.number().optional(),
    font: z.object({
      name: z.string().optional(),
      size: z.number().optional(),
      color: z.string().optional(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
    }).optional(),
  })),
  selection: z.object({
    text: z.string(),
    startParagraphIndex: z.number().optional(),
    endParagraphIndex: z.number().optional(),
  }),
});

const chatSchema = z.object({
  messages: z.array(z.object({ role: z.enum(["system", "user", "assistant", "tool"]), content: z.string().min(1), tool_call_id: z.string().optional(), name: z.string().optional() })),
  documentContext: z.string().optional(),
  documentStructure: documentStructureSchema.optional(),
  selection: z.string().optional(),
  insertMode: z.enum(["chat_only", "smart_action", "replace_selection", "append_end"]).optional(),
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

  // Build document structure description for the system prompt
  let documentStructureDescription: string | undefined;
  if (payload.documentStructure) {
    documentStructureDescription = describeDocumentStructure(payload.documentStructure);
  }

  const enrichedMessages: ChatMessage[] = contextParts.length
    ? [{ role: "system", content: contextParts.join("\n\n") }, ...messages]
    : messages;

  const insertMode = payload.insertMode || "smart_action";

  return { provider, enrichedMessages, retrieved, documentStructureDescription, insertMode };
}

// --- Non-streaming chat endpoint ---

app.post("/v1/chat", async (req, res) => {
  const parsed = chatSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: "Invalid chat payload", details: parsed.error.flatten() });
  }

  try {
    const { provider, enrichedMessages, retrieved, documentStructureDescription, insertMode } = await buildChatContext(parsed.data);
    const result = await callOpenAICompatible(
      provider,
      enrichedMessages,
      retrieved,
      WORD_TOOLS,
      insertMode,
      documentStructureDescription
    );
    return res.json({
      ok: true,
      reply: result.reply,
      actionPlan: result.actionPlan,
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

// --- Streaming chat endpoint ---

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
    const { provider, enrichedMessages, retrieved, documentStructureDescription, insertMode } = await buildChatContext(parsed.data);

    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders();

    // Send an immediate event to confirm stream startup to the client.
    res.write(`data: ${JSON.stringify({ type: "start", ts: Date.now() })}\n\n`);

    let result: { reply: string; actionPlan: ActionPlan | null };
    try {
      result = await streamOpenAICompatible(
        provider,
        enrichedMessages,
        retrieved,
        (delta) => {
          res.write(`data: ${JSON.stringify({ type: "delta", delta })}\n\n`);
        },
        abortController.signal,
        WORD_TOOLS,
        insertMode,
        documentStructureDescription
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      const shouldFallbackToNonStream = message.includes("first token not received");

      if (!shouldFallbackToNonStream) {
        throw error;
      }

      // Some OpenAI-compatible providers accept stream=true but don't emit SSE chunks.
      // Fall back to non-stream completion so client can still receive a final answer.
      result = await callOpenAICompatible(provider, enrichedMessages, retrieved, WORD_TOOLS, insertMode, documentStructureDescription);
      res.write(`data: ${JSON.stringify({ type: "fallback", reason: "no_stream_delta" })}\n\n`);
    }

    res.write(`data: ${JSON.stringify({
      type: "done",
      reply: result.reply,
      actionPlan: result.actionPlan,
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

// --- Agent Loop Endpoints ---

const agentStreamSchema = chatSchema.extend({
  enableReAct: z.boolean().optional(),
  maxIterations: z.number().int().min(1).max(10).optional(),
});

async function runAgentIteration(
  provider: ProviderConfig,
  messages: ChatMessage[],
  retrieved: RetrievalChunk[],
  insertMode: string,
  documentStructureDescription: string | undefined,
  onEvent: (event: any) => void,
  signal: AbortSignal,
  iteration: number
): Promise<{ reply: string; actionPlan: ActionPlan | null; done: boolean; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> }> {
  const result = await streamOpenAICompatible(
    provider,
    messages,
    retrieved,
    (delta) => {
      onEvent({ type: "delta", delta });
    },
    signal,
    WORD_TOOLS,
    insertMode,
    documentStructureDescription
  );

  // If LLM returned tool calls
  if (result.toolCalls && result.toolCalls.length > 0) {
    const actionPlan = parseActionPlanFromToolCalls(result.toolCalls);

    // Check if task_complete is called — stop the loop
    const taskCompleteCall = result.toolCalls.find((tc) => tc.function.name === "task_complete");
    if (taskCompleteCall) {
      let summary = "任务已完成";
      try {
        const args = JSON.parse(taskCompleteCall.function.arguments || "{}");
        summary = args.summary || summary;
      } catch { /* ignore */ }
      onEvent({ type: "task_complete", summary });
      return { reply: summary, actionPlan: null, done: true, toolCalls: result.toolCalls };
    }

    // Check if all tool calls are perception-only
    if (isPerceptionOnlyPlan(actionPlan)) {
      onEvent({ type: "tool_call", tools: result.toolCalls.map((tc) => ({ id: tc.id, tool: tc.function.name, params: JSON.parse(tc.function.arguments || "{}") })) });
      return { reply: result.reply, actionPlan, done: false, toolCalls: result.toolCalls };
    }

    // Mixed or action-only: send action_plan event and finish
    onEvent({ type: "action_plan", plan: actionPlan });
    return { reply: result.reply, actionPlan, done: true, toolCalls: result.toolCalls };
  }

  // No tool calls — final reply
  return { reply: result.reply, actionPlan: result.actionPlan, done: true };
}

app.post("/v1/chat/agent-stream", async (req, res) => {
  const parsed = agentStreamSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: "Invalid chat payload", details: parsed.error.flatten() });
  }

  const abortController = new AbortController();
  req.on("aborted", () => abortController.abort(new Error("client_aborted")));
  res.on("close", () => {
    if (!res.writableEnded && !abortController.signal.aborted) {
      abortController.abort(new Error("client_closed"));
    }
  });

  try {
    const { provider, enrichedMessages, retrieved, documentStructureDescription, insertMode } = await buildChatContext(parsed.data);
    const maxIterations = parsed.data.maxIterations ?? 5;

    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders();

    res.write(`data: ${JSON.stringify({ type: "start", ts: Date.now() })}
\n`);

    // Create session with initial messages and cache document context for agent-continue
    const sessionId = createSession(enrichedMessages);
    const createdSession = getSession(sessionId);
    if (createdSession) {
      createdSession.documentStructureDescription = documentStructureDescription;
      createdSession.insertMode = insertMode;
    }
    res.write(`data: ${JSON.stringify({ type: "session", sessionId })}
\n`);

    let currentMessages = [...enrichedMessages];
    let finalReply = "";
    let finalActionPlan: ActionPlan | null = null;
    let iteration = 0;

    while (iteration < maxIterations) {
      iteration++;

      let iterationResult: { reply: string; actionPlan: ActionPlan | null; done: boolean; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> };
      try {
        iterationResult = await runAgentIteration(
          provider,
          currentMessages,
          retrieved,
          insertMode,
          documentStructureDescription,
          (event) => res.write(`data: ${JSON.stringify(event)}
\n`),
          abortController.signal,
          iteration
        );
      } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        const shouldFallbackToNonStream = message.includes("first token not received");

        if (!shouldFallbackToNonStream) {
          throw error;
        }

        // Fallback to non-stream
        const fallbackResult = await callOpenAICompatible(
          provider,
          currentMessages,
          retrieved,
          WORD_TOOLS,
          insertMode,
          documentStructureDescription
        );

        if (fallbackResult.toolCalls && fallbackResult.toolCalls.length > 0) {
          const actionPlan = parseActionPlanFromToolCalls(fallbackResult.toolCalls);
          if (isPerceptionOnlyPlan(actionPlan)) {
            res.write(`data: ${JSON.stringify({ type: "tool_call", tools: fallbackResult.toolCalls.map((tc) => ({ id: tc.id, tool: tc.function.name, params: JSON.parse(tc.function.arguments || "{}") })) })}
\n`);
            iterationResult = { reply: fallbackResult.reply, actionPlan, done: false, toolCalls: fallbackResult.toolCalls };
          } else {
            res.write(`data: ${JSON.stringify({ type: "action_plan", plan: actionPlan })}
\n`);
            iterationResult = { reply: fallbackResult.reply, actionPlan, done: true, toolCalls: fallbackResult.toolCalls };
          }
        } else {
          res.write(`data: ${JSON.stringify({ type: "fallback", reason: "no_stream_delta" })}
\n`);
          iterationResult = { reply: fallbackResult.reply, actionPlan: fallbackResult.actionPlan, done: true };
        }
      }

      finalReply = iterationResult.reply;
      finalActionPlan = iterationResult.actionPlan;

      // Always save assistant reply to session history
      if (iterationResult.toolCalls && iterationResult.toolCalls.length > 0) {
        currentMessages.push({
          role: "assistant",
          content: iterationResult.reply,
          tool_calls: iterationResult.toolCalls,
        });
      } else if (iterationResult.actionPlan && iterationResult.actionPlan.actions.length > 0) {
        // Action plan from text parsing - create synthetic tool_calls for session
        const syntheticToolCalls = iterationResult.actionPlan.actions.map((action, index) => ({
          id: action.toolCallId || `text-parsed-${Date.now()}-${index}`,
          type: "function" as const,
          function: {
            name: action.action,
            arguments: JSON.stringify(action.params),
          },
        }));
        currentMessages.push({
          role: "assistant",
          content: iterationResult.reply,
          tool_calls: syntheticToolCalls,
        });
      } else if (iterationResult.reply) {
        currentMessages.push({
          role: "assistant",
          content: iterationResult.reply,
        });
      }

      if (iterationResult.done) {
        break;
      }

      // Wait for tool results from client
      // We must finish this response and expect client to call /v1/chat/agent-continue
      res.write(`data: ${JSON.stringify({ type: "await_tool_result", iteration })}
\n`);
      break;
    }

    // Save final state to session before ending
    const session = getSession(sessionId);
    if (session) {
      session.messages = currentMessages;
    }

    res.write(`data: ${JSON.stringify({
      type: "done",
      reply: finalReply,
      actionPlan: finalActionPlan,
      iteration,
      sessionId,
      retrievalCount: retrieved.length,
      citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })),
    })}
\n`);
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
    res.write(`data: ${JSON.stringify({ type: "error", error: message })}
\n`);
    return res.end();
  }
});

const agentContinueSchema = z.object({
  sessionId: z.string().min(1),
  toolResults: z.array(z.object({
    toolCallId: z.string(),
    toolName: z.string(),
    result: z.string(),
    success: z.boolean(),
  })),
  documentStructure: documentStructureSchema.optional(),
});

app.post("/v1/chat/agent-continue", async (req, res) => {
  const parsed = agentContinueSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: "Invalid continue payload", details: parsed.error.flatten() });
  }

  const abortController = new AbortController();
  req.on("aborted", () => abortController.abort(new Error("client_aborted")));
  res.on("close", () => {
    if (!res.writableEnded && !abortController.signal.aborted) {
      abortController.abort(new Error("client_closed"));
    }
  });

  try {
    const session = getSession(parsed.data.sessionId);
    if (!session) {
      return res.status(404).json({ error: "Session not found or expired" });
    }

    // Use fresh document structure from request if provided, otherwise fall back to session-cached context
    let documentStructureDescription = session.documentStructureDescription;
    let insertMode = session.insertMode || "smart_action";

    if (parsed.data.documentStructure) {
      documentStructureDescription = describeDocumentStructure(parsed.data.documentStructure);
      // Update session cache with fresh data
      session.documentStructureDescription = documentStructureDescription;
    }

    const { provider, retrieved } = await buildChatContext({
      messages: session.messages,
      documentContext: undefined,
      documentStructure: undefined,
      selection: undefined,
      insertMode: insertMode as any,
    });

    // Ensure the session has an assistant message with tool_calls before tool results
    // If the last message is assistant but has no tool_calls, or is not assistant,
    // we need to handle this gracefully
    const lastMsg = session.messages[session.messages.length - 1];
    if (lastMsg.role !== "assistant") {
      // The session doesn't have a proper assistant message.
      // This can happen when the action plan came from text parsing (no tool_calls).
      // In this case, we cannot continue the agent loop properly.
      // Return a meaningful error so the frontend can handle it.
      return res.status(400).json({ error: "无法继续 Agent 循环：当前会话状态不包含待处理的工具调用。操作已在本地执行完成。" });
    }

    // Append tool results to session messages
    for (const tr of parsed.data.toolResults) {
      session.messages.push({
        role: "tool",
        content: tr.result,
        tool_call_id: tr.toolCallId,
        name: tr.toolName,
      });
    }

    res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders();

    res.write(`data: ${JSON.stringify({ type: "start", ts: Date.now() })}
\n`);

    let result: { reply: string; actionPlan: ActionPlan | null; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> };
    try {
      result = await streamOpenAICompatible(
        provider,
        session.messages,
        retrieved,
        (delta) => res.write(`data: ${JSON.stringify({ type: "delta", delta })}
\n`),
        abortController.signal,
        WORD_TOOLS,
        insertMode,
        documentStructureDescription
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      const shouldFallbackToNonStream = message.includes("first token not received");

      if (!shouldFallbackToNonStream) {
        throw error;
      }

      result = await callOpenAICompatible(
        provider,
        session.messages,
        retrieved,
        WORD_TOOLS,
        insertMode,
        documentStructureDescription
      );
      res.write(`data: ${JSON.stringify({ type: "fallback", reason: "no_stream_delta" })}
\n`);
    }

    // Save the new assistant message to session for next iteration
    if (result.toolCalls && result.toolCalls.length > 0) {
      session.messages.push({
        role: "assistant",
        content: result.reply,
        tool_calls: result.toolCalls,
      });
    } else if (result.actionPlan && result.actionPlan.actions.length > 0) {
      // Action plan from text parsing - create synthetic tool_calls for session
      const syntheticToolCalls = result.actionPlan.actions.map((action, index) => ({
        id: action.toolCallId || `text-parsed-${Date.now()}-${index}`,
        type: "function" as const,
        function: {
          name: action.action,
          arguments: JSON.stringify(action.params),
        },
      }));
      session.messages.push({
        role: "assistant",
        content: result.reply,
        tool_calls: syntheticToolCalls,
      });
    } else if (result.reply) {
      session.messages.push({
        role: "assistant",
        content: result.reply,
      });
    }

    // If tool calls again, check for task_complete or send tool_call/action_plan event
    if (result.toolCalls && result.toolCalls.length > 0) {
      // Check if task_complete is called — stop the loop
      const taskCompleteCall = result.toolCalls.find((tc) => tc.function.name === "task_complete");
      if (taskCompleteCall) {
        let summary = "任务已完成";
        try {
          const args = JSON.parse(taskCompleteCall.function.arguments || "{}");
          summary = args.summary || summary;
        } catch { /* ignore */ }
        res.write(`data: ${JSON.stringify({ type: "task_complete", summary })}
\n`);
      } else {
        const actionPlan = parseActionPlanFromToolCalls(result.toolCalls);
        if (isPerceptionOnlyPlan(actionPlan)) {
          res.write(`data: ${JSON.stringify({ type: "tool_call", tools: result.toolCalls.map((tc) => ({ id: tc.id, tool: tc.function.name, params: JSON.parse(tc.function.arguments || "{}") })) })}
\n`);
        } else {
          res.write(`data: ${JSON.stringify({ type: "action_plan", plan: actionPlan })}
\n`);
        }
      }
    }

    res.write(`data: ${JSON.stringify({
      type: "done",
      reply: result.reply,
      actionPlan: result.actionPlan,
      sessionId: session.id,
      retrievalCount: retrieved.length,
      citations: retrieved.map((c) => ({ id: c.id, fileName: c.fileName })),
    })}
\n`);
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
    res.write(`data: ${JSON.stringify({ type: "error", error: message })}
\n`);
    return res.end();
  }
});

// Helper needed by agent endpoints
function parseActionPlanFromToolCalls(
  toolCalls: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }>
): ActionPlan {
  const actions = toolCalls.map((tc) => {
    let params: Record<string, any> = {};
    try {
      params = JSON.parse(tc.function.arguments);
    } catch {
      params = { raw_arguments: tc.function.arguments };
    }
    return {
      action: tc.function.name,
      params,
      description: `${tc.function.name}: ${JSON.stringify(params).slice(0, 100)}`,
    };
  });

  return {
    actions,
    explanation: actions.map((a) => a.description).join("；"),
  };
}

app.listen(port, host, () => {
  console.log(`Local agent listening on http://${host}:${port}`);
});
