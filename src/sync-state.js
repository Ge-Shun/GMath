export function normalizeOmmlFingerprint(xml) {
  return String(xml || "")
    .replace(/^\s*<\?xml[^>]*>\s*/i, "")
    .replace(/\s+xmlns(?::[\w.-]+)?="[^"]*"/g, "")
    .replace(/>\s+</g, "><")
    .replace(/\s+/g, " ")
    .trim();
}

export function makeAnchorTag() {
  const id = globalThis.crypto?.randomUUID?.()
    || `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  return `gmath:${id}`;
}

// 防抖之外再串行化 Word.run；新任务排队后，尚未开始的旧任务会自动跳过。
export class LatestTaskQueue {
  #revision = 0;
  #tail = Promise.resolve();

  enqueue(task) {
    const revision = ++this.#revision;
    const run = this.#tail.catch(() => {}).then(() => {
      if (revision !== this.#revision) return undefined;
      return task({ revision, isLatest: () => revision === this.#revision });
    });
    this.#tail = run;
    return run;
  }

  invalidate() {
    this.#revision++;
  }
}
