import type { MemoDocType, MemoTemplateRecord } from '../types/memo';

const DB_NAME = 'memo_engine_templates_db';
const DB_VERSION = 1;
const STORE_NAME = 'templates';

function openTemplatesDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('IndexedDB unavailable'));
      return;
    }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'docType' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('Failed to open template storage.'));
  });
}

export async function saveTemplateRecord(record: MemoTemplateRecord): Promise<void> {
  try {
    const db = await openTemplatesDb();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      store.put({
        ...record,
        bytes: Array.from(record.bytes),
      });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error('Failed to save template.'));
    });
    db.close();
  } catch {
    // Fallback gracefully to in-memory state if IndexedDB is blocked
  }
}

export async function deleteTemplateRecord(docType: MemoDocType): Promise<void> {
  try {
    const db = await openTemplatesDb();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).delete(docType);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error('Failed to remove template.'));
    });
    db.close();
  } catch {
    // Ignore storage errors
  }
}

export async function loadAllTemplateRecords(): Promise<Record<MemoDocType, MemoTemplateRecord | null>> {
  const empty: Record<MemoDocType, MemoTemplateRecord | null> = {
    ic: null,
    mec: null,
    board: null,
    sanction: null,
  };

  try {
    const db = await openTemplatesDb();
    const rows = await new Promise<Array<Record<string, unknown>>>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).getAll();
      req.onsuccess = () => resolve((req.result as Array<Record<string, unknown>>) || []);
      req.onerror = () => reject(req.error || new Error('Failed to load templates.'));
    });
    db.close();

    for (const row of rows) {
      const docType = row.docType as MemoDocType;
      if (docType in empty && Array.isArray(row.bytes)) {
        empty[docType] = {
          docType,
          fileName: String(row.fileName || ''),
          fileSize: Number(row.fileSize || 0),
          uploadedAt: String(row.uploadedAt || ''),
          bytes: new Uint8Array(row.bytes as number[]),
          detectedTokens: Array.isArray(row.detectedTokens) ? (row.detectedTokens as string[]) : [],
          previewSections: Array.isArray(row.previewSections)
            ? (row.previewSections as MemoTemplateRecord['previewSections'])
            : [],
        };
      }
    }
  } catch {
    // Return empty map if IndexedDB unavailable
  }

  return empty;
}
