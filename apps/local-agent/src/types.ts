export type ProviderConfig = {
  baseUrl: string;
  apiKey: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
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
