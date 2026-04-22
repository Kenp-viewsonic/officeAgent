import { RetrievalChunk } from "./types.js";

export function splitToChunks(fileName: string, content: string, size = 700): RetrievalChunk[] {
  const normalized = content.replace(/\r\n/g, "\n");
  const chunks: RetrievalChunk[] = [];

  for (let i = 0, cursor = 0; cursor < normalized.length; i += 1, cursor += size) {
    const text = normalized.slice(cursor, cursor + size).trim();
    if (!text) {
      continue;
    }
    chunks.push({
      id: `${fileName}-${i}`,
      fileName,
      text,
    });
  }

  return chunks;
}

export function keywordRetrieve(chunks: RetrievalChunk[], query: string, limit = 4): RetrievalChunk[] {
  const tokens = query
    .toLowerCase()
    .split(/\s+/)
    .map((token) => token.trim())
    .filter(Boolean);

  if (tokens.length === 0) {
    return chunks.slice(0, limit);
  }

  return [...chunks]
    .map((chunk) => {
      const lower = chunk.text.toLowerCase();
      const score = tokens.reduce((acc, token) => (lower.includes(token) ? acc + 1 : acc), 0);
      return { chunk, score };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, limit)
    .map((item) => item.chunk);
}
