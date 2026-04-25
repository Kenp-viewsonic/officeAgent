export type ProviderConfig = {
  baseUrl: string;
  apiKey: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
  firstTokenTimeout?: number;
  overallTimeout?: number;
};

export type ChatMessage = {
  role: "system" | "user" | "assistant";
  content: string;
};

export type RetrievalChunk = {
  id: string;
  fileName: string;
  text: string;
};

// --- Document Structure ---

export type DocumentParagraph = {
  index: number;
  text: string;
  style: string;
  headingLevel?: number;
  isTable: boolean;
  isList: boolean;
};

export type DocumentStructure = {
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

// --- Word Actions ---

export type WordAction = {
  action: string;
  params: Record<string, any>;
  description: string;
};

export type ActionPlan = {
  actions: WordAction[];
  explanation: string;
};

// --- Tool Definitions (OpenAI Function Calling format) ---

export type ToolFunction = {
  name: string;
  description: string;
  parameters: Record<string, any>;
};

export type ToolDefinition = {
  type: "function";
  function: ToolFunction;
};

// --- LLM Response ---

export type LlmResponse = {
  reply: string;
  actionPlan: ActionPlan | null;
};

// --- Insert Mode ---

export type InsertMode = "chat_only" | "smart_action" | "replace_selection" | "append_end";
