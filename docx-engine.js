/* Lossless Browser OOXML (.docx) Parser, Editor Model & Writer
 * Supports:
 *  - Importing any 2-column / multi-column / paragraph .docx
 *  - Native Word Threaded Comments & Multi-Round Replies (comments.xml + commentsExtended.xml + commentsIds.xml + commentsExtensible.xml)
 *  - Cyan / Green / Yellow Highlights + Auto-pairing & Manual Linking + Auto-cleanup on comment delete
 *  - Dual English Edit Modes: Track Changes (<w:del>/<w:ins> + Accept/Reject) & Direct Clean Edit
 */
(function (globalScope) {
  'use strict';

  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
  const W16CID_NS = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
  const W16CEX_NS = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
  const XML_NS = 'http://www.w3.org/XML/1998/namespace';
  const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

  const REL_COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
  const REL_COMMENTS_EX = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
  const REL_COMMENTS_IDS = 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds';
  const REL_COMMENTS_EXT = 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible';

  const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
  const CT_COMMENTS_EX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
  const CT_COMMENTS_IDS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml';
  const CT_COMMENTS_EXT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml';

  const AFTER_HL = new Set([
    'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em',
    'lang', 'eastAsianLayout', 'specVanish', 'oMath'
  ]);

  let uidCounter = 1;
  function nextUid(prefix) {
    return `${prefix}_${uidCounter++}`;
  }

  function genHexId(usedSet) {
    while (true) {
      // MS-DOCX paraId must be > 0 and < 0x80000000
      const n = Math.floor(Math.random() * 0x6FFFFFFE) + 0x10000001;
      const hex = n.toString(16).toUpperCase().padStart(8, '0');
      if (!usedSet || !usedSet.has(hex)) {
        if (usedSet) usedSet.add(hex);
        return hex;
      }
    }
  }

  function makeInitials(author) {
    if (!author) return 'RV';
    const trimmed = author.trim();
    const ascii = trimmed.replace(/[^A-Za-z0-9]/g, '');
    if (ascii.length >= 2) return ascii.slice(0, 3).toUpperCase();
    return trimmed.slice(0, 2);
  }

  function nowIsoUtc() {
    return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
  }

  function localNameOf(node) {
    if (!node || node.nodeType !== 1) return '';
    return node.localName || node.nodeName.replace(/^.*:/, '');
  }

  function getAttrLocal(el, localName) {
    if (!el || !el.attributes) return null;
    for (let i = 0; i < el.attributes.length; i++) {
      const a = el.attributes[i];
      const ln = a.localName || a.name.replace(/^.*:/, '');
      if (ln === localName) return a.value;
    }
    return null;
  }

  function parseXml(xmlStr) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) {
      throw new Error('XML 解析失败: ' + err.textContent.slice(0, 160));
    }
    return doc;
  }

  function serializeXml(doc) {
    const s = new XMLSerializer().serializeToString(doc);
    if (s.startsWith('<?xml')) return s;
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  }

  function stripHighlightAndRefFromRPr(rPrEl) {
    if (!rPrEl) return { hl: null, isCommentRef: false, baseXml: '' };
    const clone = rPrEl.cloneNode(true);
    let hl = null;
    let isCommentRef = false;
    for (const ch of Array.from(clone.childNodes)) {
      const ln = localNameOf(ch);
      if (ln === 'highlight') {
        hl = getAttrLocal(ch, 'val') || null;
        clone.removeChild(ch);
      } else if (ln === 'rStyle') {
        const val = (getAttrLocal(ch, 'val') || '').toLowerCase();
        if (val === 'commentreference' || val === 'ae') {
          isCommentRef = true;
          clone.removeChild(ch);
        }
      }
    }
    const ser = clone.childNodes.length ? new XMLSerializer().serializeToString(clone) : '';
    return { hl, isCommentRef, baseXml: ser };
  }

  function extractRunText(rEl, wantDel) {
    let txt = '';
    for (const ch of Array.from(rEl.childNodes)) {
      const ln = localNameOf(ch);
      if (!wantDel && ln === 't') txt += ch.textContent || '';
      else if (wantDel && ln === 'delText') txt += ch.textContent || '';
      else if (ln === 'tab') txt += '\t';
      else if (ln === 'br') txt += '\n';
    }
    return txt;
  }

  function tokenLen(tok) {
    if (tok.kind === 'text') return tok.text.length;
    const eff = tok.insText.length > 0 ? tok.insText : tok.delText;
    return Math.max(1, eff.length);
  }

  function tokenVisibleText(tok) {
    if (tok.kind === 'text') return tok.text;
    const eff = tok.insText.length > 0 ? tok.insText : tok.delText;
    return eff.length > 0 ? eff : '\u200B';
  }

  function paraVisibleText(pModel) {
    return pModel.tokens.map(tokenVisibleText).join('');
  }

  function coalesceParagraphTokens(pModel) {
    const out = [];
    for (const tok of pModel.tokens) {
      if (tok.kind === 'text') {
        if (!tok.text) continue;
        const prev = out[out.length - 1];
        if (prev && prev.kind === 'text' && prev.hl === tok.hl && prev.rPrBaseXml === tok.rPrBaseXml) {
          prev.text += tok.text;
        } else {
          out.push({ ...tok });
        }
      } else {
        if (!tok.delText && !tok.insText) continue;
        out.push({ ...tok });
      }
    }
    pModel.tokens = out;
  }

  /** Split pModel.tokens at exact visible offset `cut` so a token boundary exists at `cut` */
  function splitTokensAt(pModel, cut) {
    if (cut <= 0) return;
    let pos = 0;
    for (let i = 0; i < pModel.tokens.length; i++) {
      const tok = pModel.tokens[i];
      const len = tokenLen(tok);
      if (cut === pos || cut === pos + len) return;
      if (pos < cut && cut < pos + len) {
        if (tok.kind === 'text') {
          const k = cut - pos;
          const left = { ...tok, id: nextUid('tk'), text: tok.text.slice(0, k) };
          const right = { ...tok, id: nextUid('tk'), text: tok.text.slice(k) };
          pModel.tokens.splice(i, 1, left, right);
        }
        // Do not slice inside an atomic Track Change revision token
        return;
      }
      pos += len;
    }
  }

  function parseParagraphNode(pEl, rowIdx, colIdx, pIdx, rangeStarts, rangeEnds, refPoints) {
    let pPrXml = '';
    let defaultRPrXml = '';
    const tokens = [];
    let curOffset = 0;

    function processContainer(container) {
      const kids = Array.from(container.childNodes);
      for (let i = 0; i < kids.length; i++) {
        const node = kids[i];
        if (node.nodeType !== 1) continue;
        const ln = localNameOf(node);

        if (ln === 'pPr') {
          if (container === pEl) pPrXml = new XMLSerializer().serializeToString(node);
          continue;
        }
        if (ln === 'commentRangeStart') {
          const cid = getAttrLocal(node, 'id');
          if (cid !== null) rangeStarts.set(String(cid), { row: rowIdx, col: colIdx, pIdx, offset: curOffset });
          continue;
        }
        if (ln === 'commentRangeEnd') {
          const cid = getAttrLocal(node, 'id');
          if (cid !== null) rangeEnds.set(String(cid), { row: rowIdx, col: colIdx, pIdx, offset: curOffset });
          continue;
        }
        if (ln === 'r') {
          let hasCommentRef = false;
          let rPrEl = null;
          for (const rc of Array.from(node.childNodes)) {
            const rln = localNameOf(rc);
            if (rln === 'rPr') rPrEl = rc;
            else if (rln === 'commentReference') {
              hasCommentRef = true;
              const cid = getAttrLocal(rc, 'id');
              if (cid !== null) refPoints.set(String(cid), { row: rowIdx, col: colIdx, pIdx, offset: curOffset });
            }
          }
          const { hl, isCommentRef, baseXml } = stripHighlightAndRefFromRPr(rPrEl);
          if (!defaultRPrXml && baseXml && !isCommentRef && !hasCommentRef) {
            defaultRPrXml = baseXml;
          }
          const txt = extractRunText(node, false);
          if (txt.length > 0) {
            tokens.push({
              id: nextUid('tk'),
              kind: 'text',
              text: txt,
              hl: hl,
              rPrBaseXml: baseXml || defaultRPrXml
            });
            curOffset += txt.length;
          }
          continue;
        }
        if (ln === 'del' || ln === 'ins') {
          let delText = '';
          let insText = '';
          let author = getAttrLocal(node, 'author') || '审校';
          let date = getAttrLocal(node, 'date') || nowIsoUtc();
          let baseXml = defaultRPrXml;
          let hl = null;

          const collectRuns = (el, isDel) => {
            let acc = '';
            for (const c of Array.from(el.childNodes)) {
              if (localNameOf(c) === 'r') {
                const rPr = Array.from(c.childNodes).find(x => localNameOf(x) === 'rPr');
                if (rPr) {
                  const info = stripHighlightAndRefFromRPr(rPr);
                  if (info.baseXml) baseXml = info.baseXml;
                  if (info.hl) hl = info.hl;
                }
                acc += extractRunText(c, isDel);
              }
            }
            return acc;
          };

          if (ln === 'del') {
            delText = collectRuns(node, true);
            // Check if next element sibling is an immediately paired <w:ins>
            let nextIdx = i + 1;
            while (nextIdx < kids.length && kids[nextIdx].nodeType !== 1) nextIdx++;
            if (nextIdx < kids.length && localNameOf(kids[nextIdx]) === 'ins') {
              const insNode = kids[nextIdx];
              const insAuthor = getAttrLocal(insNode, 'author') || author;
              if (insAuthor === author) {
                insText = collectRuns(insNode, false);
                i = nextIdx;
              }
            }
          } else {
            insText = collectRuns(node, false);
          }

          if (delText.length > 0 || insText.length > 0) {
            const revTok = {
              id: nextUid('rev'),
              kind: 'rev',
              delText,
              insText,
              author,
              date,
              hl,
              rPrBaseXml: baseXml
            };
            tokens.push(revTok);
            curOffset += tokenLen(revTok);
          }
          continue;
        }
        // Recurse into wrappers like hyperlink, smartTag, sdtContent
        processContainer(node);
      }
    }

    processContainer(pEl);
    const pModel = { pPrXml, defaultRPrXml, tokens };
    coalesceParagraphTokens(pModel);
    return pModel;
  }

  /* Sentence splitting & monotonic alignment for readable side-by-side layout */
  function splitZhSentences(zh) {
    const spans = [];
    const n = zh.length;
    let st = 0, i = 0, qDepth = 0;
    while (i < n) {
      const ch = zh[i];
      if (ch === '“') qDepth++;
      else if (ch === '”' && qDepth > 0) qDepth--;
      else if (('。！？\n').includes(ch) && qDepth === 0) {
        let j = i + 1;
        while (j < n && ('”’）').includes(zh[j])) j++;
        if (j > st) spans.push([st, j]);
        st = j;
        i = j;
        continue;
      }
      i++;
    }
    if (st < n) spans.push([st, n]);
    return spans.length ? spans : [[0, n]];
  }

  function splitEnSentences(en) {
    const spans = [];
    const n = en.length;
    let st = 0, i = 0, inQ = false;
    while (i < n) {
      const ch = en[i];
      if (ch === '“') inQ = true;
      else if (ch === '”') inQ = false;
      else if (ch === '\n') {
        if (i + 1 > st) spans.push([st, i + 1]);
        st = i + 1;
      } else if (('.!?').includes(ch) && !inQ) {
        let j = i + 1;
        while (j < n && ('”’)"').includes(en[j])) j++;
        let k = j;
        while (k < n && en[k] === ' ') k++;
        if (k >= n || /[A-Z“]/.test(en[k])) {
          spans.push([st, k]);
          st = k;
          i = k;
          continue;
        }
      }
      i++;
    }
    if (st < n) spans.push([st, n]);
    return spans.length ? spans : [[0, n]];
  }

  function computeRowSentenceBlocks(rowModel) {
    const zhFull = getCellFullText(rowModel.cells[0]);
    const enFull = rowModel.cells[1] ? getCellFullText(rowModel.cells[1]) : '';
    if (!rowModel.cells[1] || !zhFull.length || !enFull.length) {
      return [{ label: 'P1', zhSpan: [0, zhFull.length], enSpan: [0, enFull.length] }];
    }
    const zs = splitZhSentences(zhFull);
    const es = splitEnSentences(enFull);
    const nz = zs.length, ne = es.length;
    const blocks = [];
    let z = 0;
    let ePrev = 0;
    while (z < nz) {
      let zEnd = z + 1;
      if (zEnd === nz) {
        blocks.push({
          zStart: z,
          zEnd: nz,
          label: nz - z === 1 ? `S${z + 1}` : `S${z + 1}–S${nz}`,
          zhSpan: [zs[z][0], zs[nz - 1][1]],
          enSpan: [ePrev < ne ? es[ePrev][0] : enFull.length, enFull.length]
        });
        break;
      }
      const targetFrac = zs[zEnd - 1][1] / zhFull.length;
      // Find English sentence boundary closest to targetFrac
      let bestE = ePrev;
      let bestDiff = Infinity;
      for (let cand = ePrev; cand <= ne; cand++) {
        const candPos = cand === 0 ? 0 : es[cand - 1][1];
        const diff = Math.abs((candPos / enFull.length) - targetFrac);
        if (diff < bestDiff) {
          bestDiff = diff;
          bestE = cand;
        }
      }
      // Merge short consecutive Chinese sentences if they share the same English sentence
      while (bestE <= ePrev && zEnd < nz) {
        zEnd++;
        if (zEnd === nz) { bestE = ne; break; }
        const tf = zs[zEnd - 1][1] / zhFull.length;
        for (let cand = ePrev + 1; cand <= ne; cand++) {
          const diff = Math.abs((es[cand - 1][1] / enFull.length) - tf);
          if (diff < bestDiff) { bestDiff = diff; bestE = cand; }
        }
      }
      if (bestE <= ePrev) bestE = Math.min(ne, ePrev + 1);
      const enStartChar = ePrev < ne ? es[ePrev][0] : enFull.length;
      const enEndChar = bestE > 0 ? es[Math.min(ne, bestE) - 1][1] : enStartChar;
      blocks.push({
        zStart: z,
        zEnd: zEnd,
        label: zEnd - z === 1 ? `S${z + 1}` : `S${z + 1}–S${zEnd}`,
        zhSpan: [zs[z][0], zs[zEnd - 1][1]],
        enSpan: [enStartChar, enEndChar]
      });
      ePrev = bestE;
      z = zEnd;
    }
    for (let i = 0; i < blocks.length - 1; i++) {
      blocks[i + 1].enSpan[0] = blocks[i].enSpan[1];
    }
    if (blocks.length) {
      blocks[0].zhSpan[0] = 0;
      blocks[0].enSpan[0] = 0;
      blocks[blocks.length - 1].zhSpan[1] = zhFull.length;
      blocks[blocks.length - 1].enSpan[1] = enFull.length;
    }
    return blocks;
  }

  function getCellFullText(cellModel) {
    if (!cellModel) return '';
    return cellModel.paras.map(paraVisibleText).join('\n');
  }

  function cellOffsetToParaOffset(cellModel, globalOff) {
    let acc = 0;
    for (let pIdx = 0; pIdx < cellModel.paras.length; pIdx++) {
      const pLen = paraVisibleText(cellModel.paras[pIdx]).length;
      if (globalOff <= acc + pLen || pIdx === cellModel.paras.length - 1) {
        return { pIdx, off: Math.max(0, Math.min(pLen, globalOff - acc)) };
      }
      acc += pLen + 1; // +1 for '\n'
    }
    return { pIdx: 0, off: 0 };
  }

  function paraOffsetToCellOffset(cellModel, targetPIdx, off) {
    let acc = 0;
    for (let i = 0; i < targetPIdx && i < cellModel.paras.length; i++) {
      acc += paraVisibleText(cellModel.paras[i]).length + 1;
    }
    return acc + off;
  }

  function collectCellHighlightRanges(cellModel) {
    const ranges = [];
    let base = 0;
    for (let pIdx = 0; pIdx < cellModel.paras.length; pIdx++) {
      const p = cellModel.paras[pIdx];
      let pos = 0;
      for (const tok of p.tokens) {
        const len = tokenLen(tok);
        if (tok.hl) {
          const prev = ranges[ranges.length - 1];
          if (prev && prev.pIdx === pIdx && prev.hl === tok.hl && prev.cellEnd === base + pos) {
            prev.cellEnd = base + pos + len;
            prev.end = pos + len;
          } else {
            ranges.push({
              pIdx,
              start: pos,
              end: pos + len,
              cellStart: base + pos,
              cellEnd: base + pos + len,
              hl: tok.hl
            });
          }
        }
        pos += len;
      }
      base += pos + 1;
    }
    return ranges;
  }

  /**
   * Parse a .docx ArrayBuffer into an interactive DocumentState
   */
  async function loadDocxBuffer(arrayBuffer, fileName) {
    const archive = await globalScope.ZipCodec.readZip(arrayBuffer);
    const docBytes = archive.files.get('word/document.xml');
    if (!docBytes) throw new Error('未找到 word/document.xml，不是有效的 Word 文档。');

    const docXmlStr = globalScope.ZipCodec.utf8Dec.decode(docBytes);
    const docDom = parseXml(docXmlStr);

    const rangeStarts = new Map();
    const rangeEnds = new Map();
    const refPoints = new Map();

    // Detect table rows or fallback body paragraphs
    const bodyEl = Array.from(docDom.documentElement.childNodes).find(n => localNameOf(n) === 'body');
    const tblEl = bodyEl ? Array.from(bodyEl.childNodes).find(n => localNameOf(n) === 'tbl') : null;

    const rows = [];
    if (tblEl) {
      const trEls = Array.from(tblEl.childNodes).filter(n => localNameOf(n) === 'tr');
      trEls.forEach((trEl, rIdx) => {
        const tcEls = Array.from(trEl.childNodes).filter(n => localNameOf(n) === 'tc');
        const cells = tcEls.slice(0, 2).map((tcEl, cIdx) => {
          const pEls = Array.from(tcEl.childNodes).filter(n => localNameOf(n) === 'p');
          const paras = pEls.map((pEl, pIdx) =>
            parseParagraphNode(pEl, rIdx, cIdx, pIdx, rangeStarts, rangeEnds, refPoints)
          );
          return { paras };
        });
        rows.push({ isTable: true, rowIdx: rIdx, cells });
      });
    } else if (bodyEl) {
      const pEls = Array.from(bodyEl.childNodes).filter(n => localNameOf(n) === 'p');
      const paras = pEls.map((pEl, pIdx) =>
        parseParagraphNode(pEl, 0, 0, pIdx, rangeStarts, rangeEnds, refPoints)
      );
      rows.push({ isTable: false, rowIdx: 0, cells: [{ paras }] });
    }

    // Parse word/comments.xml & word/commentsExtended.xml
    const usedHexIds = new Set();
    const rawComments = new Map(); // cid -> obj
    const paraIdToCid = new Map();
    let savedAgyLinks = null;

    if (archive.files.has('word/comments.xml')) {
      const cXmlStr = globalScope.ZipCodec.utf8Dec.decode(archive.files.get('word/comments.xml'));
      const mLinks = cXmlStr.match(/<!--agy-links:([\s\S]*?)-->/);
      if (mLinks) {
        try { savedAgyLinks = JSON.parse(mLinks[1]); } catch (e) {}
      }
      const cDom = parseXml(cXmlStr);
      for (const cEl of Array.from(cDom.documentElement.childNodes)) {
        if (localNameOf(cEl) !== 'comment') continue;
        const cid = String(getAttrLocal(cEl, 'id'));
        const author = getAttrLocal(cEl, 'author') || '审校';
        const initials = getAttrLocal(cEl, 'initials') || makeInitials(author);
        const date = getAttrLocal(cEl, 'date') || '';
        const pNodes = Array.from(cEl.childNodes).filter(n => localNameOf(n) === 'p');
        const lines = pNodes.map(pN => {
          let line = '';
          const walkT = (el) => {
            for (const k of Array.from(el.childNodes)) {
              const ln = localNameOf(k);
              if (ln === 't') line += k.textContent || '';
              else if (ln === 'br') line += '\n';
              else if (k.nodeType === 1) walkT(k);
            }
          };
          walkT(pN);
          return line;
        });
        let lastParaId = pNodes.length ? getAttrLocal(pNodes[pNodes.length - 1], 'paraId') : null;
        if (lastParaId) usedHexIds.add(lastParaId.toUpperCase());
        else lastParaId = genHexId(usedHexIds);

        rawComments.set(cid, {
          cid,
          paraId: lastParaId,
          author,
          initials,
          date,
          text: lines.join('\n').trim()
        });
        paraIdToCid.set(lastParaId, cid);
      }
    }

    // Parse extended threading info
    const parentParaMap = new Map(); // paraId -> parentParaId
    const doneParaMap = new Map();   // paraId -> bool
    if (archive.files.has('word/commentsExtended.xml')) {
      const exStr = globalScope.ZipCodec.utf8Dec.decode(archive.files.get('word/commentsExtended.xml'));
      const exDom = parseXml(exStr);
      for (const exEl of Array.from(exDom.documentElement.childNodes)) {
        if (localNameOf(exEl) !== 'commentEx') continue;
        const pid = getAttrLocal(exEl, 'paraId');
        const ppid = getAttrLocal(exEl, 'paraIdParent');
        const done = getAttrLocal(exEl, 'done') === '1';
        if (pid) {
          usedHexIds.add(pid.toUpperCase());
          doneParaMap.set(pid, done);
          if (ppid) parentParaMap.set(pid, ppid);
        }
      }
    }

    // Build root threads & nested replies
    const rootMap = new Map(); // rootCid -> thread
    const findRootCid = (cid) => {
      const visited = new Set();
      let cur = cid;
      while (cur && !visited.has(cur)) {
        visited.add(cur);
        const cObj = rawComments.get(cur);
        if (!cObj) break;
        const pPara = parentParaMap.get(cObj.paraId);
        if (!pPara || !paraIdToCid.has(pPara)) return cur;
        cur = paraIdToCid.get(pPara);
      }
      return cid;
    };

    for (const [cid, cObj] of rawComments.entries()) {
      const rCid = findRootCid(cid);
      if (rCid === cid) {
        const st = rangeStarts.get(cid) || refPoints.get(cid) || { row: 0, col: 0, pIdx: 0, offset: 0 };
        const en = rangeEnds.get(cid) || st;
        const sameP = (st.row === en.row && st.col === en.col && st.pIdx === en.pIdx);
        const anchor = {
          row: st.row,
          col: st.col,
          pIdx: st.pIdx,
          start: Math.min(st.offset, sameP ? en.offset : st.offset),
          end: sameP ? Math.max(st.offset, en.offset) : paraVisibleText(rows[st.row].cells[st.col].paras[st.pIdx]).length
        };
        rootMap.set(cid, {
          id: nextUid('th'),
          cid,
          paraId: cObj.paraId,
          author: cObj.author,
          initials: cObj.initials,
          date: cObj.date,
          text: cObj.text,
          done: !!doneParaMap.get(cObj.paraId),
          anchor,
          linkedSpans: [],
          replies: []
        });
      }
    }

    for (const [cid, cObj] of rawComments.entries()) {
      const rCid = findRootCid(cid);
      if (rCid !== cid && rootMap.has(rCid)) {
        rootMap.get(rCid).replies.push({
          id: nextUid('rp'),
          cid,
          paraId: cObj.paraId,
          author: cObj.author,
          initials: cObj.initials,
          date: cObj.date,
          text: cObj.text
        });
      }
    }

    const threads = Array.from(rootMap.values());
    sortThreadsInPlace(threads, rows);

    // Compute sentence blocks per row
    rows.forEach(r => { r.blocks = computeRowSentenceBlocks(r); });

    // Restore saved linkedSpans OR smart-pair English highlights by aligned sentence block
    if (savedAgyLinks && typeof savedAgyLinks === 'object') {
      for (const th of threads) {
        if (Array.isArray(savedAgyLinks[th.cid])) {
          th.linkedSpans = savedAgyLinks[th.cid];
        }
      }
    } else {
      autoLinkHighlightsOnImport(threads, rows);
    }

    return {
      fileName: fileName || '未命名文稿.docx',
      archive,
      docXmlStr,
      rows,
      threads,
      usedHexIds
    };
  }

  function autoLinkHighlightsOnImport(threads, rows) {
    rows.forEach((rObj, rIdx) => {
      if (!rObj.cells[1]) return;
      const enRanges = collectCellHighlightRanges(rObj.cells[1]);
      if (!enRanges.length) return;

      rObj.blocks.forEach((blk) => {
        const blkThreads = threads.filter(th => {
          if (th.anchor.row !== rIdx || th.anchor.col !== 0) return false;
          const cOff = paraOffsetToCellOffset(rObj.cells[0], th.anchor.pIdx, th.anchor.start);
          return cOff >= blk.zhSpan[0] && cOff < blk.zhSpan[1];
        });
        if (!blkThreads.length) return;

        const blkCyan = enRanges.filter(rg =>
          (rg.hl === 'cyan' || rg.hl === 'yellow') &&
          rg.cellStart >= blk.enSpan[0] && rg.cellStart < blk.enSpan[1]
        );
        const blkGreen = enRanges.filter(rg =>
          rg.hl === 'green' &&
          rg.cellStart >= blk.enSpan[0] && rg.cellStart < blk.enSpan[1]
        );

        if (blkThreads.length === 1) {
          blkThreads[0].linkedSpans.push(
            ...blkCyan.map(rg => ({ row: rIdx, col: 1, pIdx: rg.pIdx, start: rg.start, end: rg.end, color: rg.hl })),
            ...blkGreen.map(rg => ({ row: rIdx, col: 1, pIdx: rg.pIdx, start: rg.start, end: rg.end, color: 'green' }))
          );
        } else if (blkCyan.length > 0) {
          // Pair monotonically by relative position in block
          blkCyan.forEach((rg, idx) => {
            const targetIdx = Math.min(blkThreads.length - 1, Math.floor((idx * blkThreads.length) / blkCyan.length));
            blkThreads[targetIdx].linkedSpans.push({
              row: rIdx, col: 1, pIdx: rg.pIdx, start: rg.start, end: rg.end, color: rg.hl
            });
          });
        }
      });
    });
  }

  function sortThreadsInPlace(threads, rows) {
    threads.sort((a, b) => {
      if (a.anchor.row !== b.anchor.row) return a.anchor.row - b.anchor.row;
      if (a.anchor.col !== b.anchor.col) return a.anchor.col - b.anchor.col;
      if (a.anchor.pIdx !== b.anchor.pIdx) return a.anchor.pIdx - b.anchor.pIdx;
      if (a.anchor.start !== b.anchor.start) return a.anchor.start - b.anchor.start;
      return a.anchor.end - b.anchor.end;
    });
  }

  /* ==================== EDITING PRIMITIVES ==================== */

  function rebaseSpansInPara(docState, row, col, pIdx, fromA, fromB, newLen) {
    const delta = newLen - (fromB - fromA);
    if (delta === 0) return;
    const adjust = (sp) => {
      if (sp.row !== row || sp.col !== col || sp.pIdx !== pIdx) return;
      if (sp.end <= fromA) return;
      if (sp.start >= fromB) {
        sp.start += delta;
        sp.end += delta;
      } else {
        sp.start = Math.min(sp.start, fromA);
        sp.end = Math.max(fromA + newLen, sp.end + delta);
      }
    };
    for (const th of docState.threads) {
      adjust(th.anchor);
      for (const lk of th.linkedSpans) adjust(lk);
    }
  }

  /** Apply or clear highlight color (`'cyan'|'green'|'yellow'|null`) on [start, end) of (row, col, pIdx) */
  function setParaHighlight(docState, row, col, pIdx, start, end, color) {
    if (end <= start) return;
    const p = docState.rows[row].cells[col].paras[pIdx];
    splitTokensAt(p, start);
    splitTokensAt(p, end);
    let pos = 0;
    for (const tok of p.tokens) {
      const len = tokenLen(tok);
      if (pos >= start && pos + len <= end) {
        tok.hl = color || null;
      }
      pos += len;
    }
    coalesceParagraphTokens(p);
  }

  /** Replace [start, end) in (row, col, pIdx) either via Track Changes or Direct Edit */
  function editParagraphRange(docState, row, col, pIdx, start, end, newText, mode, author) {
    const p = docState.rows[row].cells[col].paras[pIdx];
    splitTokensAt(p, start);
    splitTokensAt(p, end);

    let firstIdx = -1, lastIdx = -1, pos = 0;
    let oldPlain = '';
    let sampleRPr = p.defaultRPrXml;
    let sampleHl = null;

    for (let i = 0; i < p.tokens.length; i++) {
      const tok = p.tokens[i];
      const len = tokenLen(tok);
      if (pos >= start && pos + len <= end) {
        if (firstIdx < 0) firstIdx = i;
        lastIdx = i;
        if (tok.rPrBaseXml) sampleRPr = tok.rPrBaseXml;
        if (tok.hl) sampleHl = tok.hl;
        oldPlain += tok.kind === 'text' ? tok.text : (tok.insText || tok.delText);
      } else if (pos === start && start === end && firstIdx < 0) {
        firstIdx = i;
        lastIdx = i - 1;
        if (tok.rPrBaseXml) sampleRPr = tok.rPrBaseXml;
      }
      pos += len;
    }
    if (firstIdx < 0) {
      firstIdx = p.tokens.length;
      lastIdx = p.tokens.length - 1;
    }

    const removeCount = Math.max(0, lastIdx - firstIdx + 1);
    let replacement = [];
    let newVisibleLen = 0;

    if (mode === 'track') {
      if (oldPlain === newText) return;
      const revTok = {
        id: nextUid('rev'),
        kind: 'rev',
        delText: oldPlain,
        insText: newText,
        author: author || '审校',
        date: nowIsoUtc(),
        hl: sampleHl,
        rPrBaseXml: sampleRPr
      };
      replacement = [revTok];
      newVisibleLen = tokenLen(revTok);
    } else {
      if (newText.length > 0) {
        replacement = [{
          id: nextUid('tk'),
          kind: 'text',
          text: newText,
          hl: sampleHl,
          rPrBaseXml: sampleRPr
        }];
      }
      newVisibleLen = newText.length;
    }

    p.tokens.splice(firstIdx, removeCount, ...replacement);
    rebaseSpansInPara(docState, row, col, pIdx, start, end, newVisibleLen);
    coalesceParagraphTokens(p);
    docState.rows[row].blocks = computeRowSentenceBlocks(docState.rows[row]);
  }

  /** Accept or Reject a single Track Change revision token by its token.id */
  function resolveTrackToken(docState, row, col, pIdx, revTokenId, accept) {
    const p = docState.rows[row].cells[col].paras[pIdx];
    let pos = 0;
    for (let i = 0; i < p.tokens.length; i++) {
      const tok = p.tokens[i];
      const oldLen = tokenLen(tok);
      if (tok.id === revTokenId && tok.kind === 'rev') {
        const finalStr = accept ? tok.insText : tok.delText;
        if (finalStr.length > 0) {
          p.tokens.splice(i, 1, {
            id: nextUid('tk'),
            kind: 'text',
            text: finalStr,
            hl: tok.hl,
            rPrBaseXml: tok.rPrBaseXml
          });
        } else {
          p.tokens.splice(i, 1);
        }
        rebaseSpansInPara(docState, row, col, pIdx, pos, pos + oldLen, finalStr.length);
        coalesceParagraphTokens(p);
        docState.rows[row].blocks = computeRowSentenceBlocks(docState.rows[row]);
        return true;
      }
      pos += oldLen;
    }
    return false;
  }

  /* ==================== LOSSLESS .DOCX SERIALIZATION ==================== */

  function createRunElement(docDom, rPrBaseXml, hlColor, textContent, isDelText) {
    const r = docDom.createElementNS(W_NS, 'w:r');
    let rPr = null;
    if (rPrBaseXml) {
      try {
        const parsed = parseXml(rPrBaseXml).documentElement;
        rPr = docDom.importNode(parsed, true);
      } catch (e) {}
    }
    if (hlColor) {
      if (!rPr) rPr = docDom.createElementNS(W_NS, 'w:rPr');
      const hlEl = docDom.createElementNS(W_NS, 'w:highlight');
      hlEl.setAttributeNS(W_NS, 'w:val', hlColor);
      let inserted = false;
      for (const ch of Array.from(rPr.childNodes)) {
        if (AFTER_HL.has(localNameOf(ch))) {
          rPr.insertBefore(hlEl, ch);
          inserted = true;
          break;
        }
      }
      if (!inserted) rPr.appendChild(hlEl);
    }
    if (rPr && rPr.childNodes.length > 0) {
      r.appendChild(rPr);
    }

    const parts = textContent.split('\n');
    parts.forEach((seg, idx) => {
      if (idx > 0) r.appendChild(docDom.createElementNS(W_NS, 'w:br'));
      if (seg.length > 0 || parts.length === 1) {
        const t = docDom.createElementNS(W_NS, isDelText ? 'w:delText' : 'w:t');
        t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
        t.textContent = seg;
        r.appendChild(t);
      }
    });
    return r;
  }

  function createCommentRefRun(docDom, cid) {
    const r = docDom.createElementNS(W_NS, 'w:r');
    const rPr = docDom.createElementNS(W_NS, 'w:rPr');
    const rStyle = docDom.createElementNS(W_NS, 'w:rStyle');
    rStyle.setAttributeNS(W_NS, 'w:val', 'CommentReference');
    rPr.appendChild(rStyle);
    r.appendChild(rPr);
    const cr = docDom.createElementNS(W_NS, 'w:commentReference');
    cr.setAttributeNS(W_NS, 'w:id', String(cid));
    r.appendChild(cr);
    return r;
  }

  function rebuildParagraphDom(docDom, origPEl, pModel, paraThreads, revIdCounter) {
    // Work on a deep clone of pModel so in-memory UI tokens stay clean
    const workP = {
      pPrXml: pModel.pPrXml,
      defaultRPrXml: pModel.defaultRPrXml,
      tokens: pModel.tokens.map(t => ({ ...t }))
    };

    for (const th of paraThreads) {
      splitTokensAt(workP, th.anchor.start);
      splitTokensAt(workP, th.anchor.end);
    }

    // Clear origPEl children, re-attach pPr
    while (origPEl.firstChild) origPEl.removeChild(origPEl.firstChild);
    if (workP.pPrXml) {
      try {
        const pPrNode = docDom.importNode(parseXml(workP.pPrXml).documentElement, true);
        origPEl.appendChild(pPrNode);
      } catch (e) {}
    }

    const startsAt = new Map();
    const endsAt = new Map();
    for (const th of paraThreads) {
      if (!startsAt.has(th.anchor.start)) startsAt.set(th.anchor.start, []);
      startsAt.get(th.anchor.start).push(th);
      if (!endsAt.has(th.anchor.end)) endsAt.set(th.anchor.end, []);
      endsAt.get(th.anchor.end).push(th);
    }

    const emitStarts = (offset) => {
      for (const th of (startsAt.get(offset) || [])) {
        const allIds = [th.exportCid, ...th.replies.map(rp => rp.exportCid)];
        for (const id of allIds) {
          const rs = docDom.createElementNS(W_NS, 'w:commentRangeStart');
          rs.setAttributeNS(W_NS, 'w:id', String(id));
          origPEl.appendChild(rs);
        }
      }
    };

    const emitEnds = (offset) => {
      for (const th of (endsAt.get(offset) || [])) {
        const allIds = [th.exportCid, ...th.replies.map(rp => rp.exportCid)];
        for (const id of allIds) {
          const re = docDom.createElementNS(W_NS, 'w:commentRangeEnd');
          re.setAttributeNS(W_NS, 'w:id', String(id));
          origPEl.appendChild(re);
          origPEl.appendChild(createCommentRefRun(docDom, id));
        }
      }
    };

    let pos = 0;
    emitStarts(0);
    // Point-comments at offset 0 (start === 0 && end === 0)
    if (endsAt.has(0)) {
      const zeroPoint = endsAt.get(0).filter(th => th.anchor.start === 0);
      for (const th of zeroPoint) {
        const allIds = [th.exportCid, ...th.replies.map(rp => rp.exportCid)];
        for (const id of allIds) {
          const re = docDom.createElementNS(W_NS, 'w:commentRangeEnd');
          re.setAttributeNS(W_NS, 'w:id', String(id));
          origPEl.appendChild(re);
          origPEl.appendChild(createCommentRefRun(docDom, id));
        }
      }
      endsAt.set(0, endsAt.get(0).filter(th => th.anchor.start !== 0));
    }

    for (const tok of workP.tokens) {
      const len = tokenLen(tok);
      if (tok.kind === 'text') {
        origPEl.appendChild(createRunElement(docDom, tok.rPrBaseXml || workP.defaultRPrXml, tok.hl, tok.text, false));
      } else if (tok.kind === 'rev') {
        if (tok.delText.length > 0) {
          const delEl = docDom.createElementNS(W_NS, 'w:del');
          delEl.setAttributeNS(W_NS, 'w:id', String(revIdCounter.val++));
          delEl.setAttributeNS(W_NS, 'w:author', tok.author || '审校');
          delEl.setAttributeNS(W_NS, 'w:date', tok.date || nowIsoUtc());
          delEl.appendChild(createRunElement(docDom, tok.rPrBaseXml || workP.defaultRPrXml, tok.hl, tok.delText, true));
          origPEl.appendChild(delEl);
        }
        if (tok.insText.length > 0) {
          const insEl = docDom.createElementNS(W_NS, 'w:ins');
          insEl.setAttributeNS(W_NS, 'w:id', String(revIdCounter.val++));
          insEl.setAttributeNS(W_NS, 'w:author', tok.author || '审校');
          insEl.setAttributeNS(W_NS, 'w:date', tok.date || nowIsoUtc());
          insEl.appendChild(createRunElement(docDom, tok.rPrBaseXml || workP.defaultRPrXml, tok.hl, tok.insText, false));
          origPEl.appendChild(insEl);
        }
      }
      pos += len;
      emitEnds(pos);
      emitStarts(pos);
    }
  }

  function escXml(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  async function exportDocxBytes(docState) {
    const docDom = parseXml(docState.docXmlStr);
    sortThreadsInPlace(docState.threads, docState.rows);

    // Assign clean sequential numeric IDs for export
    let nextCid = 1;
    for (const th of docState.threads) {
      th.exportCid = String(nextCid++);
      if (!th.paraId) th.paraId = genHexId(docState.usedHexIds);
      for (const rp of th.replies) {
        rp.exportCid = String(nextCid++);
        if (!rp.paraId) rp.paraId = genHexId(docState.usedHexIds);
      }
    }

    const revIdCounter = { val: 9001 };
    const bodyEl = Array.from(docDom.documentElement.childNodes).find(n => localNameOf(n) === 'body');
    const tblEl = bodyEl ? Array.from(bodyEl.childNodes).find(n => localNameOf(n) === 'tbl') : null;

    if (tblEl && docState.rows[0]?.isTable) {
      const trEls = Array.from(tblEl.childNodes).filter(n => localNameOf(n) === 'tr');
      trEls.forEach((trEl, rIdx) => {
        const rModel = docState.rows[rIdx];
        if (!rModel) return;
        const tcEls = Array.from(trEl.childNodes).filter(n => localNameOf(n) === 'tc');
        tcEls.slice(0, 2).forEach((tcEl, cIdx) => {
          const cModel = rModel.cells[cIdx];
          if (!cModel) return;
          const pEls = Array.from(tcEl.childNodes).filter(n => localNameOf(n) === 'p');
          pEls.forEach((pEl, pIdx) => {
            const pThreads = docState.threads.filter(
              th => th.anchor.row === rIdx && th.anchor.col === cIdx && th.anchor.pIdx === pIdx
            );
            rebuildParagraphDom(docDom, pEl, cModel.paras[pIdx], pThreads, revIdCounter);
          });
        });
      });
    } else if (bodyEl && docState.rows[0]) {
      const pEls = Array.from(bodyEl.childNodes).filter(n => localNameOf(n) === 'p');
      pEls.forEach((pEl, pIdx) => {
        const pThreads = docState.threads.filter(
          th => th.anchor.row === 0 && th.anchor.col === 0 && th.anchor.pIdx === pIdx
        );
        rebuildParagraphDom(docDom, pEl, docState.rows[0].cells[0].paras[pIdx], pThreads, revIdCounter);
      });
    }

    const outFiles = new Map(docState.archive.files);
    outFiles.set('word/document.xml', serializeXml(docDom));

    // Build word/comments.xml + word/commentsExtended.xml + word/commentsIds.xml + word/commentsExtensible.xml
    const linksMeta = {};
    const commentXmlBlocks = [];
    const exXmlBlocks = [];
    const cidXmlBlocks = [];
    const cextXmlBlocks = [];

    const appendCommentRecord = (cid, paraId, parentParaId, done, author, initials, dateStr, textStr) => {
      const lines = (textStr || '').split('\n');
      const pXmls = lines.map((ln, idx) => {
        const pid = (idx === lines.length - 1) ? paraId : genHexId(docState.usedHexIds);
        return `<w:p w14:paraId="${pid}" w14:textId="77777777"><w:r><w:t xml:space="preserve">${escXml(ln)}</w:t></w:r></w:p>`;
      });
      const dt = dateStr || nowIsoUtc();
      commentXmlBlocks.push(
        `<w:comment w:id="${escXml(cid)}" w:author="${escXml(author || '审校')}" w:date="${escXml(dt)}" w:initials="${escXml(initials || makeInitials(author))}">${pXmls.join('')}</w:comment>`
      );
      const parentAttr = parentParaId ? ` w15:paraIdParent="${parentParaId}"` : '';
      exXmlBlocks.push(`<w15:commentEx w15:paraId="${paraId}"${parentAttr} w15:done="${done ? '1' : '0'}"/>`);
      const durableId = genHexId(docState.usedHexIds);
      cidXmlBlocks.push(`<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="${durableId}"/>`);
      cextXmlBlocks.push(`<w16cex:commentExtensible w16cex:durableId="${durableId}" w16cex:dateUtc="${escXml(dt)}"/>`);
    };

    for (const th of docState.threads) {
      linksMeta[th.exportCid] = th.linkedSpans;
      appendCommentRecord(th.exportCid, th.paraId, null, th.done, th.author, th.initials, th.date, th.text);
      for (const rp of th.replies) {
        appendCommentRecord(rp.exportCid, rp.paraId, th.paraId, th.done, rp.author, rp.initials, rp.date, rp.text);
      }
    }

    const safeLinksJson = JSON.stringify(linksMeta).replace(/--/g, '\\u002d\\u002d');
    const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">` +
      `<!--agy-links:${safeLinksJson}-->` +
      `${commentXmlBlocks.join('')}</w:comments>`;

    const commentsExXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<w15:commentsEx xmlns:w15="${W15_NS}" xmlns:w="${W_NS}">${exXmlBlocks.join('')}</w15:commentsEx>`;

    const commentsIdsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<w16cid:commentsIds xmlns:w16cid="${W16CID_NS}">${cidXmlBlocks.join('')}</w16cid:commentsIds>`;

    const commentsExtXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<w16cex:commentsExtensible xmlns:w16cex="${W16CEX_NS}">${cextXmlBlocks.join('')}</w16cex:commentsExtensible>`;

    outFiles.set('word/comments.xml', commentsXml);
    outFiles.set('word/commentsExtended.xml', commentsExXml);
    outFiles.set('word/commentsIds.xml', commentsIdsXml);
    outFiles.set('word/commentsExtensible.xml', commentsExtXml);

    // Ensure relationships in word/_rels/document.xml.rels
    if (outFiles.has('word/_rels/document.xml.rels')) {
      const relDom = parseXml(globalScope.ZipCodec.utf8Dec.decode(outFiles.get('word/_rels/document.xml.rels')));
      const relRoot = relDom.documentElement;
      const existingTypes = new Set();
      const usedRIds = new Set();
      for (const el of Array.from(relRoot.childNodes)) {
        if (localNameOf(el) === 'Relationship') {
          existingTypes.add(getAttrLocal(el, 'Type'));
          usedRIds.add(getAttrLocal(el, 'Id'));
        }
      }
      const ensureRel = (typeUri, target) => {
        if (existingTypes.has(typeUri)) return;
        let n = 60;
        while (usedRIds.has(`rId${n}`)) n++;
        const rid = `rId${n}`;
        usedRIds.add(rid);
        const rel = relDom.createElementNS(REL_NS, 'Relationship');
        rel.setAttribute('Id', rid);
        rel.setAttribute('Type', typeUri);
        rel.setAttribute('Target', target);
        relRoot.appendChild(rel);
      };
      ensureRel(REL_COMMENTS, 'comments.xml');
      ensureRel(REL_COMMENTS_EX, 'commentsExtended.xml');
      ensureRel(REL_COMMENTS_IDS, 'commentsIds.xml');
      ensureRel(REL_COMMENTS_EXT, 'commentsExtensible.xml');
      outFiles.set('word/_rels/document.xml.rels', serializeXml(relDom));
    }

    // Ensure overrides in [Content_Types].xml
    if (outFiles.has('[Content_Types].xml')) {
      const ctDom = parseXml(globalScope.ZipCodec.utf8Dec.decode(outFiles.get('[Content_Types].xml')));
      const ctRoot = ctDom.documentElement;
      const existingParts = new Set();
      for (const el of Array.from(ctRoot.childNodes)) {
        if (localNameOf(el) === 'Override') {
          existingParts.add(getAttrLocal(el, 'PartName'));
        }
      }
      const ensureCt = (partName, contentType) => {
        if (existingParts.has(partName)) return;
        const ov = ctDom.createElementNS(CT_NS, 'Override');
        ov.setAttribute('PartName', partName);
        ov.setAttribute('ContentType', contentType);
        ctRoot.appendChild(ov);
      };
      ensureCt('/word/comments.xml', CT_COMMENTS);
      ensureCt('/word/commentsExtended.xml', CT_COMMENTS_EX);
      ensureCt('/word/commentsIds.xml', CT_COMMENTS_IDS);
      ensureCt('/word/commentsExtensible.xml', CT_COMMENTS_EXT);
      outFiles.set('[Content_Types].xml', serializeXml(ctDom));
    }

    return globalScope.ZipCodec.writeZip({ order: docState.archive.order, files: outFiles });
  }

  /* Deep snapshot of mutable state for Undo/Redo */
  function snapshotMutableState(docState) {
    return JSON.stringify({
      rows: docState.rows.map(r => ({
        isTable: r.isTable,
        rowIdx: r.rowIdx,
        cells: r.cells.map(c => ({
          paras: c.paras.map(p => ({
            pPrXml: p.pPrXml,
            defaultRPrXml: p.defaultRPrXml,
            tokens: p.tokens.map(t => ({ ...t }))
          }))
        }))
      })),
      threads: docState.threads.map(th => ({
        ...th,
        anchor: { ...th.anchor },
        linkedSpans: th.linkedSpans.map(s => ({ ...s })),
        replies: th.replies.map(rp => ({ ...rp }))
      }))
    });
  }

  function restoreMutableState(docState, snapJson) {
    const data = JSON.parse(snapJson);
    docState.rows = data.rows;
    docState.rows.forEach(r => { r.blocks = computeRowSentenceBlocks(r); });
    docState.threads = data.threads;
  }

  globalScope.DocxEngine = {
    loadDocxBuffer,
    exportDocxBytes,
    getCellFullText,
    paraVisibleText,
    tokenLen,
    tokenVisibleText,
    cellOffsetToParaOffset,
    paraOffsetToCellOffset,
    setParaHighlight,
    editParagraphRange,
    resolveTrackToken,
    sortThreadsInPlace,
    computeRowSentenceBlocks,
    snapshotMutableState,
    restoreMutableState,
    genHexId,
    makeInitials,
    nowIsoUtc,
    nextUid
  };
})(typeof window !== 'undefined' ? window : globalThis);
