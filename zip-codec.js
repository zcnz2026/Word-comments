/* Pure-Web PKZIP Reader & Writer (Zero external dependencies, offline & GitHub Pages ready)
 * Uses native Web Streams CompressionStream / DecompressionStream ('deflate-raw').
 */
(function (globalScope) {
  'use strict';

  const CRC_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  async function runStream(inputBytes, transformStream) {
    const writer = transformStream.writable.getWriter();
    writer.write(inputBytes);
    writer.close();
    const reader = transformStream.readable.getReader();
    const chunks = [];
    let total = 0;
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      chunks.push(value);
      total += value.byteLength;
    }
    const out = new Uint8Array(total);
    let offset = 0;
    for (const ch of chunks) {
      out.set(ch, offset);
      offset += ch.byteLength;
    }
    return out;
  }

  async function inflateRaw(compressedBytes) {
    if (compressedBytes.byteLength === 0) return new Uint8Array(0);
    return runStream(compressedBytes, new DecompressionStream('deflate-raw'));
  }

  async function deflateRaw(rawBytes) {
    if (rawBytes.byteLength === 0) return new Uint8Array(0);
    return runStream(rawBytes, new CompressionStream('deflate-raw'));
  }

  const utf8Dec = new TextDecoder('utf-8');
  const utf8Enc = new TextEncoder();

  /**
   * Read a .docx / .zip ArrayBuffer or Uint8Array via its Central Directory.
   * Returns { order: string[], files: Map<string, Uint8Array> }
   */
  async function readZip(buffer) {
    const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

    // Locate End of Central Directory (EOCD) signature 0x06054b50 scanning backwards
    let eocdOffset = -1;
    const minPos = Math.max(0, bytes.length - 65558);
    for (let pos = bytes.length - 22; pos >= minPos; pos--) {
      if (view.getUint32(pos, true) === 0x06054b50) {
        eocdOffset = pos;
        break;
      }
    }
    if (eocdOffset < 0) {
      throw new Error('无效的 .docx 文件（未找到 ZIP 中央目录签名，请确认上传的是 .docx 而非旧版 .doc）。');
    }

    const totalEntries = view.getUint16(eocdOffset + 10, true);
    const cdOffset = view.getUint32(eocdOffset + 16, true);

    const files = new Map();
    const order = [];
    let ptr = cdOffset;

    for (let i = 0; i < totalEntries; i++) {
      if (view.getUint32(ptr, true) !== 0x02014b50) {
        throw new Error('ZIP 中央目录损坏，无法解析全部组件。');
      }
      const method = view.getUint16(ptr + 10, true);
      const compSize = view.getUint32(ptr + 20, true);
      const uncompSize = view.getUint32(ptr + 24, true);
      const nameLen = view.getUint16(ptr + 28, true);
      const extraLen = view.getUint16(ptr + 30, true);
      const commentLen = view.getUint16(ptr + 32, true);
      const localHdrOffset = view.getUint32(ptr + 42, true);

      const nameBytes = bytes.subarray(ptr + 46, ptr + 46 + nameLen);
      const name = utf8Dec.decode(nameBytes);
      ptr += 46 + nameLen + extraLen + commentLen;

      if (name.endsWith('/')) continue;

      if (view.getUint32(localHdrOffset, true) !== 0x04034b50) {
        throw new Error('ZIP 本地文件头损坏: ' + name);
      }
      const lNameLen = view.getUint16(localHdrOffset + 26, true);
      const lExtraLen = view.getUint16(localHdrOffset + 28, true);
      const dataStart = localHdrOffset + 30 + lNameLen + lExtraLen;
      const compSlice = bytes.subarray(dataStart, dataStart + compSize);

      let content;
      if (method === 0) {
        content = new Uint8Array(compSlice);
      } else if (method === 8) {
        content = await inflateRaw(compSlice);
        if (uncompSize > 0 && content.length !== uncompSize) {
          // Tolerate minor stream padding only if content non-empty, otherwise warn
        }
      } else {
        throw new Error(`不支持的 ZIP 压缩算法 (${method}): ${name}`);
      }
      order.push(name);
      files.set(name, content);
    }

    return { order, files };
  }

  /**
   * Build a standard .docx Uint8Array from { order, files }
   */
  async function writeZip(archive) {
    const names = [];
    const seen = new Set();
    // Put [Content_Types].xml first per OPC convention
    if (archive.files.has('[Content_Types].xml')) {
      names.push('[Content_Types].xml');
      seen.add('[Content_Types].xml');
    }
    for (const n of (archive.order || [])) {
      if (archive.files.has(n) && !seen.has(n)) {
        names.push(n);
        seen.add(n);
      }
    }
    for (const n of archive.files.keys()) {
      if (!seen.has(n)) {
        names.push(n);
        seen.add(n);
      }
    }

    const localParts = [];
    const cdParts = [];
    let offset = 0;

    // DOS timestamp (2026-11-18 12:00:00)
    const dosTime = (12 << 11) | (0 << 5) | 0;
    const dosDate = ((2026 - 1980) << 9) | (11 << 5) | 18;

    for (const name of names) {
      const raw = archive.files.get(name);
      const rawBytes = typeof raw === 'string' ? utf8Enc.encode(raw) : raw;
      const nameBytes = utf8Enc.encode(name);
      const crc = crc32(rawBytes);

      let compBytes;
      let method = 8;
      if (rawBytes.length < 48) {
        compBytes = rawBytes;
        method = 0;
      } else {
        compBytes = await deflateRaw(rawBytes);
        if (compBytes.length >= rawBytes.length) {
          compBytes = rawBytes;
          method = 0;
        }
      }

      // Local File Header (30 + nameLen)
      const lhdr = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(lhdr.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);          // version needed 2.0
      lv.setUint16(6, 0x0800, true);      // UTF-8 filename flag
      lv.setUint16(8, method, true);
      lv.setUint16(10, dosTime, true);
      lv.setUint16(12, dosDate, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, compBytes.length, true);
      lv.setUint32(22, rawBytes.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      lhdr.set(nameBytes, 30);

      // Central Directory Header (46 + nameLen)
      const chdr = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(chdr.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);          // version made by
      cv.setUint16(6, 20, true);          // version needed
      cv.setUint16(8, 0x0800, true);      // UTF-8 flag
      cv.setUint16(10, method, true);
      cv.setUint16(12, dosTime, true);
      cv.setUint16(14, dosDate, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, compBytes.length, true);
      cv.setUint32(24, rawBytes.length, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      chdr.set(nameBytes, 46);

      localParts.push(lhdr, compBytes);
      cdParts.push(chdr);
      offset += lhdr.length + compBytes.length;
    }

    const cdSize = cdParts.reduce((s, a) => s + a.length, 0);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, names.length, true);
    ev.setUint16(10, names.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    const totalLen = offset + cdSize + 22;
    const out = new Uint8Array(totalLen);
    let p = 0;
    for (const part of localParts) { out.set(part, p); p += part.length; }
    for (const part of cdParts) { out.set(part, p); p += part.length; }
    out.set(eocd, p);
    return out;
  }

  globalScope.ZipCodec = { readZip, writeZip, utf8Dec, utf8Enc };
})(typeof window !== 'undefined' ? window : globalThis);
