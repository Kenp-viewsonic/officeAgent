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
  // --- 感知类工具 (Perception) ---
  {
    type: "function",
    function: {
      name: "read_document",
      description: "动态读取文档中指定范围或类型的段落内容。用于在操作前确认目标位置、验证插入结果、或获取超出初始上下文的详细信息。建议在不确定段落索引或标题位置时先调用此工具确认。",
      parameters: {
        type: "object",
        properties: {
          mode: {
            type: "string",
            enum: ["paragraph_range", "heading_context", "selection", "cursor_surrounding"],
            description: "读取模式：paragraph_range=按段落范围读取；heading_context=读取指定标题及其下所有子内容直到下一个同级标题；selection=读取当前选区；cursor_surrounding=读取光标前后内容",
          },
          paragraph_index: { type: "number", description: "mode=paragraph_range 时的起始段落序号（从0开始）" },
          count: { type: "number", description: "mode=paragraph_range 时读取的段落数量，默认5" },
          heading_text: { type: "string", description: "mode=heading_context 时，要查找的标题文本" },
          surrounding_chars: { type: "number", description: "mode=cursor_surrounding 时，光标前后各读取多少字符，默认500" },
        },
        required: ["mode"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_selection_info",
      description: "获取当前选区的精确信息，包括选区文本、起始/结束段落序号、是否仅为光标位置（无选中文本）。用于判断应使用 insert_at_cursor 还是 replace_selection。",
      parameters: {
        type: "object",
        properties: {},
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_document_stats",
      description: "获取文档统计信息：总段落数、总字符数、各级标题列表及其段落序号。用于快速了解文档整体结构。",
      parameters: {
        type: "object",
        properties: {},
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_paragraph_format",
      description: "获取指定段落的详细格式信息，包括段落内每个文本片段（run）的字体、字号、加粗、斜体、颜色。当需要精确了解段落内的内联格式变化时使用（例如：段落中部分文字使用了不同字体或加粗）。返回 paragraphFont（段落默认字体）和 runs 数组（每个 run 有独立字体信息）。如果 runs 只有 1 个，说明全段格式统一。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "要查看格式的段落序号（从0开始）" },
        },
        required: ["paragraph_index"],
      },
    },
  },

  // --- 操作类工具 (Action) ---
  {
    type: "function",
    function: {
      name: "insert_after_heading",
      description: "在指定标题后插入新内容。通过 heading_text 定位标题，支持指定格式。如果标题匹配失败，会插入到文档末尾。",
      parameters: {
        type: "object",
        properties: {
          heading_text: { type: "string", description: "要查找的标题文本" },
          content: { type: "string", description: "要插入的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式（如 <b>加粗</b>、<i>斜体</i>、<span style='font-family:黑体;font-size:14pt'>指定字体</span>）" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal。注意：content_format='html' 时此参数仍可用于设置段落样式",
          },
        },
        required: ["heading_text", "content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_at_cursor",
      description: "在当前光标位置插入内容。如果当前有选区，选区内容会被替换。这是最常用的插入工具，语义明确为「在光标处插入」。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "要插入的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式（如 <b>加粗</b>、<i>斜体</i>、<span style='font-family:黑体;font-size:14pt'>指定字体</span>）" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
          },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "replace_selection",
      description: "替换当前选中的文本内容。如果当前没有选区（仅光标），行为等同于 insert_at_cursor。建议先调用 get_selection_info 确认选区状态。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "替换后的新内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
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
          content: { type: "string", description: "要追加的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
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
      description: "在文档开头插入内容。注意：空文档时可能需要特殊处理。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "要插入的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
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
      description: "在指定段落序号后插入内容。段落序号从文档结构中获取。注意：插入/删除操作后段落索引会变化，建议先调用 read_document 确认索引。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "目标段落的序号（从文档结构中获取，从0开始）" },
          content: { type: "string", description: "要插入的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
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
      description: "删除指定段落序号的内容。建议先调用 read_document 确认段落内容后再删除。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "要删除的段落序号（从0开始）" },
        },
        required: ["paragraph_index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_and_replace",
      description: "在文档中查找文本并替换为新文本。仅替换第一处匹配，大小写敏感。如需更多控制，请使用 find_and_replace_v2。",
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
      name: "find_and_replace_v2",
      description: "增强版查找替换。支持全文档替换、大小写控制、全词匹配，并返回替换次数。",
      parameters: {
        type: "object",
        properties: {
          find_text: { type: "string", description: "要查找的文本" },
          replace_text: { type: "string", description: "替换后的文本" },
          replace_all: { type: "boolean", description: "是否替换所有匹配项，默认 false（仅替换第一处）" },
          match_case: { type: "boolean", description: "是否区分大小写，默认 false" },
          match_whole_word: { type: "boolean", description: "是否全词匹配，默认 false" },
        },
        required: ["find_text", "replace_text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "apply_rich_format",
      description: "对指定段落或选区应用富文本格式（字体、颜色、加粗、斜体、超链接）。支持对段落内特定文字片段（通过 text_to_format 精确匹配）单独设置格式，实现段落内混合排版。",
      parameters: {
        type: "object",
        properties: {
          target_mode: {
            type: "string",
            enum: ["selection", "paragraph_index", "last_inserted"],
            description: "目标：selection=当前选区；paragraph_index=指定段落；last_inserted=上次插入的段落",
          },
          paragraph_index: { type: "number", description: "target_mode=paragraph_index 时的段落序号" },
          text_to_format: { type: "string", description: "段落内要格式化的精确文本片段。指定后只对该文本应用格式，而非整个段落/选区。例如段落中'重点内容'三个字需要加粗，则填'重点内容'" },
          font: {
            type: "object",
            properties: {
              name: { type: "string", description: "字体名称，如'微软雅黑'" },
              size: { type: "number", description: "字号（磅），如 14" },
              color: { type: "string", description: "字体颜色（十六进制），如 #FF0000" },
              bold: { type: "boolean", description: "是否加粗" },
              italic: { type: "boolean", description: "是否斜体" },
            },
          },
          hyperlink: {
            type: "object",
            properties: {
              text: { type: "string", description: "超链接显示文本" },
              url: { type: "string", description: "超链接URL" },
            },
          },
        },
        required: ["target_mode"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "replace_paragraph",
      description: "替换指定段落的全部文本内容。用于修改现有段落文字，保留段落格式。建议先调用 read_document 确认段落内容后再替换。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "要替换的段落序号（从0开始）" },
          content: { type: "string", description: "替换后的新内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
        },
        required: ["paragraph_index", "content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "set_paragraph_style",
      description: "修改指定段落的样式（如标题级别、列表格式）。用于调整已有段落的格式而不改变内容。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "目标段落序号（从0开始）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "要设置的格式",
          },
        },
        required: ["paragraph_index", "format"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "merge_paragraphs",
      description: "将两个相邻段落合并为一个段落。第二个段落的内容会追加到第一个段落的末尾。",
      parameters: {
        type: "object",
        properties: {
          first_paragraph_index: { type: "number", description: "第一个段落的序号（合并后内容保留在此位置）" },
          second_paragraph_index: { type: "number", description: "第二个段落的序号（合并后此段落会被删除）" },
          separator: { type: "string", description: "合并时两段之间的分隔符，默认为空格", default: " " },
        },
        required: ["first_paragraph_index", "second_paragraph_index"],
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
  {
    type: "function",
    function: {
      name: "task_complete",
      description: "标记任务已完成，停止 Agent 循环。当所有操作都已执行完毕，或者任务不需要进一步操作时调用此工具。",
      parameters: {
        type: "object",
        properties: {
          summary: { type: "string", description: "任务完成的总结说明" },
        },
        required: ["summary"],
      },
    },
  },
];

// --- Build document structure description ---

function describeFontBrief(font?: { name?: string; size?: number; color?: string; bold?: boolean; italic?: boolean }): string {
  if (!font) return "";
  const parts: string[] = [];
  if (font.name) parts.push(font.name);
  if (font.size) parts.push(`${font.size}pt`);
  if (font.bold) parts.push("加粗");
  if (font.italic) parts.push("斜体");
  if (font.color && !["#000000", "#000000ff", "#000"].includes(font.color.toLowerCase())) parts.push(font.color);
  return parts.length > 0 ? ` [${parts.join(" ")}]` : "";
}

export function describeDocumentStructure(structure: {
  totalParagraphs: number;
  totalCharacters?: number;
  paragraphs: Array<{ index: number; text: string; style: string; headingLevel?: number; isList: boolean; charCount?: number; font?: { name?: string; size?: number; color?: string; bold?: boolean; italic?: boolean } }>;
  selection: { text: string; startParagraphIndex?: number; endParagraphIndex?: number };
}): string {
  const lines: string[] = [];
  lines.push(`文档共 ${structure.totalParagraphs} 段${structure.totalCharacters ? `，约 ${structure.totalCharacters} 字符` : ""}。结构概览：`);

  for (const p of structure.paragraphs) {
    const charInfo = p.charCount ? `(${p.charCount}字)` : "";
    const fontInfo = describeFontBrief(p.font);
    if (p.headingLevel) {
      lines.push(`  ${"#".repeat(p.headingLevel)} [段落${p.index}]${charInfo}${fontInfo} ${p.text}`);
    } else if (p.isList) {
      lines.push(`  - [段落${p.index}]${charInfo}${fontInfo} ${p.text}`);
    } else {
      lines.push(`  [段落${p.index}]${charInfo}${fontInfo} ${p.text.slice(0, 100)}`);
    }
  }

  if (structure.selection.text) {
    const selInfo = structure.selection.startParagraphIndex !== undefined
      ? `（位于段落 ${structure.selection.startParagraphIndex}${structure.selection.endParagraphIndex !== structure.selection.startParagraphIndex ? ` ~ ${structure.selection.endParagraphIndex}` : ""}）`
      : "";
    lines.push(`\n当前选区${selInfo}：${structure.selection.text.slice(0, 500)}`);
  } else {
    lines.push(`\n当前无选中文本（光标模式）`);
  }

  lines.push(`\n注意：以上字体为段落级别默认字体。Word 支持段落内内联格式变化（如部分文字使用不同字体、加粗等），如需查看某段落的详细格式分布，请使用 get_paragraph_format 工具。`);

  return lines.join("\n");
}

// --- Build system prompt based on mode and document structure ---

function buildSystemPrompt(
  hasTools: boolean,
  insertMode: string,
  documentStructureDescription?: string
): ChatMessage {
  let content = `你是一个面向 Word 文档编辑的智能 Agent。你的工作是理解用户意图，通过「感知 → 思考 → 操作」的循环来完成任务。

## 核心规则
1. 根据文档结构和用户意图选择最合适的工具和参数
2. 如果用户要求在某个位置插入内容，优先使用 insert_after_heading 或 insert_after_paragraph，而不是让用户手动定位
3. 如果用户要求删除内容，使用 delete_paragraph 或 find_and_replace_v2
4. 如果用户只是提问或需要建议，使用 reply_only
5. 当使用 insert/replace 工具时，content 参数应该是可直接写入 Word 的纯文本，不要使用 Markdown 标记
6. format 参数用于指定插入内容的格式，默认为 normal
7. 在引用知识库时标注来源编号

## 感知优先原则
- 如果不确定文档当前状态、目标位置、或索引是否有效，先调用 read_document 或 get_selection_info 确认
- 插入/删除操作后段落索引会变化，后续操作应基于新的索引
- 如果 find_and_replace_v2 返回 replaced=0，说明查找失败，不要继续基于该假设执行后续操作
- content 参数必须是非空字符串
- 注意：段落级别字体信息只是该段落的默认/主字体。Word 文档支持在同一段落内对不同文字片段设置不同格式（内联格式），如需精确了解段落内的格式分布（例如某段落中部分文字加粗、部分使用不同字体），请使用 get_paragraph_format 工具获取逐片段的格式详情`;

  if (hasTools) {
    content += `

## 可用工具

### 感知类工具（用于确认文档状态，建议在操作前调用）
- read_document: 动态读取文档片段（按段落范围、标题上下文、选区、光标周围）
- get_selection_info: 获取当前选区精确信息（文本、段落索引、是否仅为光标）
- get_document_stats: 获取文档统计信息（总段落数、字符数、标题列表）
- get_paragraph_format: 获取指定段落的详细格式信息（段落内每个文本片段的字体、字号、加粗、斜体、颜色），用于检测内联格式变化

### 操作类工具（用于修改文档）
- insert_after_heading: 在指定标题后插入新内容
- insert_at_cursor: 在当前光标位置插入内容（推荐，语义最明确）
- replace_selection: 替换当前选中的文本
- replace_paragraph: 替换指定段落的全部文本内容（保留格式）
- insert_at_end: 在文档末尾追加内容
- insert_at_start: 在文档开头插入内容
- insert_after_paragraph: 在指定段落序号后插入内容
- delete_paragraph: 删除指定段落
- find_and_replace: 查找并替换文本（简单版，仅替换第一处）
- find_and_replace_v2: 增强版查找替换（支持全文档替换、大小写控制、返回替换次数）
- set_paragraph_style: 修改指定段落的样式（标题级别、列表格式）
- apply_rich_format: 对段落或选区应用富文本格式（字体、颜色、加粗、超链接）
- merge_paragraphs: 合并两个相邻段落
- reply_only: 仅回复文本，不执行文档操作
- task_complete: 标记任务已完成，停止 Agent 循环

### 工具使用指南
- 定位内容：优先使用 find_and_replace_v2 或 read_document(mode="heading_context")，而非依赖可能过时的 paragraph_index
- 插入内容：光标处插入 → insert_at_cursor；标题后插入 → insert_after_heading；文末 → insert_at_end；文首 → insert_at_start
- 修改内容：替换整段 → replace_paragraph；修改文字 → find_and_replace_v2；改样式 → set_paragraph_style
- 格式控制：format 参数只控制段落级别样式（标题、列表等）。插入工具支持 content_format="html"，可直接在 content 中使用 HTML 标签实现内联格式（如 <b>加粗</b>、<i>斜体</i>、<span style="font-family:黑体;font-size:14pt">指定字体</span>）。对已有文本的局部格式化，使用 apply_rich_format 配合 text_to_format 参数精确匹配目标文字
- 混合格式插入示例：content_format="html", content="这是<b>重点内容</b>和<i>斜体内容</i>的混合段落"
- 不确定时：先 read_document 确认位置和内容，再执行操作

### 任务完成规则
- 当你确认所有操作都已执行完毕，或者任务不需要进一步操作时，必须调用 task_complete 工具来结束 Agent 循环
- task_complete 的 summary 参数应包含任务完成的总结说明
- 不要无限循环调用感知工具，如果已经获取了足够的信息就执行操作并调用 task_complete

请根据用户意图选择合适的工具调用。如果用户只是提问，使用 reply_only。`;
  }

  if (insertMode === "chat_only") {
    content += `

当前模式为「仅对话」，请只使用 reply_only 工具回复用户，不要执行任何文档操作。`;
  } else if (insertMode === "replace_selection") {
    content += `

当前模式为「替换选区」，请优先使用 replace_selection 或 insert_at_cursor 工具。`;
  } else if (insertMode === "append_end") {
    content += `

当前模式为「追加到文末」，请优先使用 insert_at_end 工具。`;
  } else if (insertMode === "smart_action") {
    content += `

当前模式为「智能操作」，请根据用户意图自主选择最合适的工具。建议在不确定时先调用感知类工具确认文档状态。`;
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
      toolCallId: tc.id,
    });
  }

  return {
    actions,
    explanation: actions.map((a) => a.description).join("；"),
  };
}

function describeAction(actionName: string, params: Record<string, any>): string {
  const fmtSuffix = params.format && params.format !== "normal" ? `（格式: ${params.format}）` : "";
  switch (actionName) {
    // 感知类工具
    case "read_document":
      if (params.mode === "paragraph_range") return `读取段落 ${params.paragraph_index ?? 0} 起共 ${params.count ?? 5} 段`;
      if (params.mode === "heading_context") return `读取标题"${params.heading_text}"及其子内容`;
      if (params.mode === "selection") return "读取当前选区内容";
      if (params.mode === "cursor_surrounding") return `读取光标周围 ${params.surrounding_chars ?? 500} 字符`;
      return `读取文档（模式: ${params.mode}）`;
    case "get_selection_info":
      return "获取当前选区信息";
    case "get_document_stats":
      return "获取文档统计信息";
    case "get_paragraph_format":
      return `获取段落${params.paragraph_index}的详细格式`;
    // 操作类工具
    case "insert_after_heading":
      return `在标题"${params.heading_text}"后插入${params.content_format === "html" ? "HTML格式" : ""}内容${fmtSuffix}`;
    case "insert_at_cursor":
      return `在光标处插入${params.content_format === "html" ? "HTML格式" : ""}内容${fmtSuffix}`;
    case "replace_selection":
      return `替换选区${params.content_format === "html" ? "HTML格式" : ""}内容${fmtSuffix}`;
    case "insert_at_end":
      return `追加${params.content_format === "html" ? "HTML格式" : ""}内容到文档末尾${fmtSuffix}`;
    case "insert_at_start":
      return `插入${params.content_format === "html" ? "HTML格式" : ""}内容到文档开头${fmtSuffix}`;
    case "insert_after_paragraph":
      return `在段落${params.paragraph_index}后插入${params.content_format === "html" ? "HTML格式" : ""}内容${fmtSuffix}`;
    case "delete_paragraph":
      return `删除段落${params.paragraph_index}`;
    case "find_and_replace":
      return `将"${params.find_text}"替换为"${params.replace_text}"`;
    case "find_and_replace_v2":
      return `查找替换"${params.find_text}"→"${params.replace_text}"${params.replace_all ? "（全部）" : "（首处）"}${params.match_case ? " 区分大小写" : ""}`;
    case "apply_rich_format":
      if (params.hyperlink) return `对${params.target_mode}应用超链接"${params.hyperlink.text}"`;
      if (params.text_to_format) return `对段落内文本"${params.text_to_format}"应用富文本格式`;
      return `对${params.target_mode}应用富文本格式`;
    case "replace_paragraph":
      return `替换段落${params.paragraph_index}的内容${params.content_format === "html" ? "（HTML格式）" : ""}`;
    case "set_paragraph_style":
      return `将段落${params.paragraph_index}设置为${params.format}格式`;
    case "merge_paragraphs":
      return `合并段落${params.first_paragraph_index}和段落${params.second_paragraph_index}${params.separator ? `（分隔符: "${params.separator}"）` : ""}`;
    case "reply_only":
      return "仅回复文本";
    case "task_complete":
      return `任务完成：${params.summary || ""}`;
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

    const timestamp = Date.now();
    const actions: WordAction[] = plan.actions.map((a: any, index: number) => ({
      action: a.action || a.name || "reply_only",
      params: a.params || a.parameters || {},
      description: a.description || describeAction(a.action || a.name || "reply_only", a.params || a.parameters || {}),
      toolCallId: `text-parsed-${timestamp}-${index}`,
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
): Promise<{ reply: string; actionPlan: ActionPlan | null; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> }> {
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
    return { reply, actionPlan, toolCalls: message.tool_calls };
  }

  // No tool_calls — check for text-based action plan (fallback)
  const textContent = message.content?.trim() || "模型没有返回可用内容。";
  const textActionPlan = parseActionPlanFromText(textContent);
  const reply = textActionPlan ? extractTextReply(textContent) || textActionPlan.explanation : textContent;

  return { reply, actionPlan: textActionPlan };
}

// Helper to check if an action plan contains only perception tools
export function isPerceptionOnlyPlan(plan: ActionPlan): boolean {
  const perceptionTools = ["read_document", "get_selection_info", "get_document_stats", "get_paragraph_format"];
  return plan.actions.length > 0 && plan.actions.every((a) => perceptionTools.includes(a.action));
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
): Promise<{ reply: string; actionPlan: ActionPlan | null; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> }> {
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
      return { reply, actionPlan, toolCalls: message.tool_calls };
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
    return { reply, actionPlan, toolCalls };
  }

  // No tool calls — check for text-based action plan
  const textActionPlan = parseActionPlanFromText(fullText);
  const reply = textActionPlan ? extractTextReply(fullText) || textActionPlan.explanation : fullText.trim() || "模型没有返回可用内容。";
  return { reply, actionPlan: textActionPlan };
}
