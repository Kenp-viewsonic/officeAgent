# Word Agent 工具与工作流分析

> 本文档基于当前代码（`apps/local-agent/src/llm.ts`、`apps/word-addin/src/main.ts`）整理，供 LLM 和开发者参考。

---

## 一、当前提供给 LLM 的工具（OpenAI Function Calling）

| # | 工具名 | 说明 | 参数 | 常见问题 |
|---|--------|------|------|----------|
| 1 | `insert_after_heading` | 在指定标题后插入内容 | `heading_text`(string), `content`(string), `format`(enum) | 标题匹配失败时 fallback 到文档末尾，但用户无感知 |
| 2 | `replace_selection` | 替换当前选区 | `content`(string), `format`(enum) | 若选区为空则变成在光标处插入，行为可能不符合预期 |
| 3 | `insert_at_end` | 在文档末尾追加 | `content`(string), `format`(enum) | 相对安全，但若文档为空，`getLast()` 可能抛错 |
| 4 | `insert_at_start` | **在文档开头插入** | `content`(string), `format`(enum) | **高危**：空文档时 `getFirst()` 报 `InvalidArgument`；空 content 也会报错 |
| 5 | `insert_after_paragraph` | 在指定段落序号后插入 | `paragraph_index`(number), `content`(string), `format`(enum) | 段落序号由前端结构提供，LLM 可能误用索引 |
| 6 | `delete_paragraph` | 删除指定段落 | `paragraph_index`(number) | 同上，索引误用风险 |
| 7 | `find_and_replace` | 查找并替换 | `find_text`(string), `replace_text`(string) | 仅替换第一处匹配；`matchCase: true` 导致大小写敏感，可能找不到 |
| 8 | `reply_only` | 仅回复文本，不操作文档 | `content`(string) | 无风险 |

### format 枚举值
```
"normal" | "heading1" | "heading2" | "heading3" | "bullet_list" | "numbered_list"
```

### 工具设计的结构性缺陷

1. **没有 `"insert_at_cursor"` 工具**  
   用户直觉是在当前光标位置插入，但现有工具只能 `"replace_selection"` 或 `"insert_at_end"`。若选区为空，`replace_selection` 行为等同于在光标处插入，但 LLM 无法判断当前是否有选区。

2. **没有 `"get_document_state"` 或 `"read_range"` 工具**  
   LLM 仅能看到前端传来的静态 `documentStructure`（最多 80 段，每段截断 200 字），无法在执行工具链中间读取最新文档状态。这导致多步操作（如"在 A 后面插入，再在 B 后面插入"）可能基于过时的段落索引失败。

3. **`insert_after_paragraph` 依赖前端截断后的索引**  
   前端只传前 80 段，如果文档超过 80 段，LLM 调用 `insert_after_paragraph` 时引用的索引可能指向错误位置或 fallback 到末尾。

4. **`find_and_replace` 能力太弱**  
   - 只替换第一处
   - `matchCase: true`（严格大小写）
   - 不支持正则
   - 没有返回"替换了几处"的信息给 LLM

5. **缺少 undo / 原子性保证**  
   每个工具独立执行，如果 ActionPlan 第 2 步失败，第 1 步不会回滚。用户看到的是半成品文档。

---

## 二、Word 编辑工作流分析

### 2.1 当前数据流

```
[用户输入] → [前端 main.ts]
                │
                ├─ 读取文档结构 (getStructuredContext) ──→ Word JS API
                │
                ├─ 发送 POST /v1/chat/stream ──→ [local-agent server.ts]
                │                                      │
                │                                      ├─ 构建 system prompt
                │                                      ├─ 检索知识库 (keywordRetrieve)
                │                                      └─ 调用 LLM (streamOpenAICompatible)
                │                                         (tools=WORD_TOOLS, tool_choice=auto)
                │
                └─ 接收 SSE stream ──→ 解析出 actionPlan
                                           │
                                           ▼
                                    [ActionPlan 预览面板]
                                           │
                                    [用户点击"确认"]
                                           │
                                           ▼
                                    [前端 executeActionPlan]
                                           │
                                    逐个调用 executeAction
                                           │
                                    Word.run(...) → Word JS API
```

### 2.2 关键问题点

#### A. LLM 的决策信息不足

当前 `system prompt` 包含：
- 文档上下文（截断到 3000 字）
- 文档结构（最多 80 段，每段截断 200 字）
- 当前选区文本
- 知识库检索片段

**缺失信息**：
- 光标位置（`selection.startParagraphIndex` 字段有定义但前端没有真正填入）
- 文档总字数精确值（虽然有 `totalCharacters`）
- 各段落的实际字符数（前端截断了）
- 用户上一次操作的结果（LLM 不知道上一步是否成功）

#### B. 工具调用 vs. 简单插入的割裂

前端有两种完全不同的插入路径：

| 路径 | 触发条件 | 执行方式 |
|------|----------|----------|
| **ActionPlan 路径** | `insertMode === "smart_action"` 且 LLM 返回 `tool_calls` | 先预览，再逐个 `executeAction` |
| **简单插入路径** | `insertMode === "replace_selection" / "append_end"` 或无 actionPlan | `showSimplePreview` → `confirmSimplePreview` |

**问题**：
- 即使用户选择了 `"smart_action"`，如果 LLM 没有返回 tool_calls，回退到简单路径时不会自动执行，而是仅显示文本。
- `"insert_start"` 简单路径被映射到了 `insertLastReplyAtCursor`（函数名和实际行为不符，实际是 `"insert_start"`）。

#### C. 错误处理过于粗糙

- `executeAction` 里的 `try/catch` 只捕获到 `Error.message`，Office.js 的错误对象包含 `code`、`traceMessages`、`debugInfo` 等大量有用信息，之前被完全丢弃。
- 用户看到的只有：`❌ 插入内容到文档开头: InvalidArgument`，无法判断是空文档、空内容、还是 Word API 其他限制。

#### D. Markdown 清洗的双刃剑

`cleanupMarkdownForWord` 会去除：
- 代码块标记
- 标题符号 `#`
- 加粗/斜体 `**`、`*`、`_`
- 列表符号
- 引用块 `>`
- 链接 `[text](url)` → `text (url)`

**问题**：
- 如果 LLM 返回的内容被清洗后变成空字符串，Word API 会报 `InvalidArgument`。
- 链接转换可能产生不自然的文本。
- 没有保留列表结构（只是简单替换为 `- `）。

---

## 三、工作流改进建议

### 3.1 短期（增强鲁棒性）

1. **空文档保护**：所有使用 `getFirst()` / `getLast()` 的地方都加 fallback 到 `body.insertParagraph(..., "Start")`。
2. **空内容保护**：`executeAction` 入口处统一校验 `content` 非空。
3. **错误信息丰富化**：保留 Office.js 错误的完整字段（已部分实现）。
4. **修复 `insert_after_paragraph` 越界行为**：当前越界时静默 fallback 到末尾，应该明确报错或询问用户。

### 3.2 中期（工具重构）

1. **新增 `insert_at_cursor` 工具**  
   明确替代 `"replace_selection"` 在空选区时的模糊语义。

2. **新增 `read_document` 工具**  
   让 LLM 能在工具链中间读取指定段落或范围，而不是依赖一次性的静态结构。

3. **改造 `find_and_replace`**  
   - 支持 `replace_all` 参数
   - 支持 `matchCase` 参数
   - 返回替换次数

4. **引入事务/原子性概念**  
   ActionPlan 执行前可以先把所有操作在内存中验证一遍（或至少标记依赖关系），失败时提供 rollback 方案。

### 3.3 长期（交互升级）

1. **LLM 能感知操作结果**  
   每个 action 执行后把结果（成功/失败、实际插入位置）反馈给 LLM，支持多轮工具调用。

2. **用户介入点细化**  
   不只是"确认/取消"整个 Plan，而是允许用户：
   - 修改某一步的参数
   - 跳过某一步
   - 在某一步后暂停询问

3. **段落索引的稳定性**  
   为每段生成稳定的 ID（如基于内容哈希），而不是依赖易变的数组索引。

---

## 四、附录：前端 ↔ Agent API 契约

### 4.1 前端发送的 Chat Payload

```typescript
{
  messages: ChatMessage[];           // 用户输入 + 历史
  documentContext: string;           // 截断到 3000 字的文档全文
  documentStructure?: DocumentStructure; // 结构化段落信息
  selection: string;                 // 当前选区文本
  insertMode: "chat_only" | "smart_action" | "replace_selection" | "append_end";
}
```

### 4.2 Agent 返回的 Stream Event

```typescript
type StreamEvent =
  | { type: "start"; ts: number }
  | { type: "delta"; delta: string }
  | { type: "fallback"; reason: string }  // 流式失败回退到非流式
  | { type: "done"; reply: string; actionPlan?: ActionPlan | null; retrievalCount?: number; citations?: Array<{fileName}> }
  | { type: "error"; error: string };
```

### 4.3 前端 WordAction 执行契约

```typescript
type WordAction = {
  action: string;           // 工具名，如 "insert_at_start"
  params: Record<string, any>;
  description: string;      // 人类可读描述，用于预览面板
};
```

---

## 五、导致 "InvalidArgument" 的具体根因汇总

| 场景 | 根因 | 代码位置 |
|------|------|----------|
| 空文档 + insert_at_start | `body.paragraphs.getFirst()` 在空文档上抛 InvalidArgument | `main.ts: executeAction` (insert_at_start) |
| 空文档 + insert_at_end | `body.paragraphs.getLast()` 在空文档上抛 InvalidArgument | `main.ts: executeAction` (insert_at_end) |
| content 为空字符串 | `insertParagraph("", "Before")` 被 Word API 拒绝 | 所有 insert/replace 工具 |
| content 仅含空白 | `cleanupMarkdownForWord` 清洗后可能只剩空格/换行 | `cleanupMarkdownForWord` |
| firstPara.text 为 undefined | 未正确 load("text") 就读取 | `getStructuredContext` 中的 sync 顺序 |
