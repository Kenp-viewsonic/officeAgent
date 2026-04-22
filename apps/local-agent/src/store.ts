import fs from "node:fs/promises";
import path from "node:path";
import { ProviderConfig, RetrievalChunk } from "./types.js";

const dataRoot = path.resolve(process.cwd(), "data");
const configFile = path.join(dataRoot, "provider-config.json");
const kbFile = path.join(dataRoot, "kb-chunks.json");

async function ensureDataRoot() {
  await fs.mkdir(dataRoot, { recursive: true });
}

export async function saveProviderConfig(config: ProviderConfig): Promise<void> {
  await ensureDataRoot();
  await fs.writeFile(configFile, JSON.stringify(config, null, 2), "utf-8");
}

export async function loadProviderConfig(): Promise<ProviderConfig | null> {
  try {
    const raw = await fs.readFile(configFile, "utf-8");
    return JSON.parse(raw) as ProviderConfig;
  } catch {
    return null;
  }
}

export async function appendChunks(chunks: RetrievalChunk[]): Promise<void> {
  await ensureDataRoot();
  const current = await loadChunks();
  const merged = [...current, ...chunks];
  await fs.writeFile(kbFile, JSON.stringify(merged, null, 2), "utf-8");
}

export async function loadChunks(): Promise<RetrievalChunk[]> {
  try {
    const raw = await fs.readFile(kbFile, "utf-8");
    return JSON.parse(raw) as RetrievalChunk[];
  } catch {
    return [];
  }
}

export async function getKbStats(): Promise<{
  chunkCount: number;
  fileCount: number;
  files: string[];
}> {
  const chunks = await loadChunks();
  const fileSet = new Set(chunks.map((chunk) => chunk.fileName));
  return {
    chunkCount: chunks.length,
    fileCount: fileSet.size,
    files: [...fileSet].sort((a, b) => a.localeCompare(b)).slice(-20),
  };
}
