/* Word 译文批注审阅工作台 — Interactive UI Controller */
(function () {
  'use strict';

  const Engine = window.DocxEngine;

  let docState = null;
  let undoStack = [];
  let redoStack = [];

  const ui = {
    viewMode: localStorage.getItem('agy_view_mode') || 'split',     // 'split' | 'inline'
    editMode: localStorage.getItem('agy_edit_mode') || 'track',     // 'track' | 'direct'
    scrollSync: localStorage.getItem('agy_scroll_sync') !== '0',    // default ON
    reviewer: localStorage.getItem('agy_reviewer') || '审校',
    statusFilter: 'ALL',   // 'ALL' | 'OPEN' | 'DONE' | 'REPLIED'
    authorFilter: 'ALL',
    searchQuery: '',
    activeThreadId: null,
    currentSelection: null, // { row, col, pIdx, start, end, text, rect }
    pendingModalAction: null
  };

  let pointerDownInfo = null;
  let pointerInDocPane = false;
  let suppressDocScrollUntil = 0;
  let scrollSyncTimer = null;
  let activeScrollAnim = new WeakMap();

  function smoothScrollElementTo(el, targetTop, durationMs) {
    if (!el) return;
    const maxScroll = Math.max(0, el.scrollHeight - el.clientHeight);
    const clampedTarget = Math.max(0, Math.min(maxScroll, targetTop));
    if (window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
      el.scrollTop = clampedTarget;
      return;
    }
    const prevAnim = activeScrollAnim.get(el);
    if (prevAnim) cancelAnimationFrame(prevAnim);

    const startTop = el.scrollTop;
    const diff = clampedTarget - startTop;
    if (Math.abs(diff) < 2) {
      el.scrollTop = clampedTarget;
      return;
    }
    const startTime = performance.now();
    const dur = durationMs || 260;

    function step(now) {
      const t = Math.min(1, (now - startTime) / dur);
      const ease = 1 - Math.pow(1 - t, 3); // cubic ease-out
      el.scrollTop = startTop + diff * ease;
      if (t < 1) {
        activeScrollAnim.set(el, requestAnimationFrame(step));
      } else {
        activeScrollAnim.delete(el);
      }
    }
    activeScrollAnim.set(el, requestAnimationFrame(step));
  }

  /* ==================== INIT & DIALOG FALLBACKS ==================== */
  function init() {
    const revInput = document.getElementById('reviewerInput');
    revInput.value = ui.reviewer;
    revInput.addEventListener('input', () => {
      ui.reviewer = revInput.value.trim() || '审校';
      localStorage.setItem('agy_reviewer', ui.reviewer);
    });

    syncModeButtons();

    // File input & global drag-drop
    const fileInput = document.getElementById('docxFileInput');
    fileInput.addEventListener('change', async (ev) => {
      const f = ev.target.files && ev.target.files[0];
      if (f) await openDocxFile(f);
      fileInput.value = '';
    });

    const dropCard = document.getElementById('dropCard');
    dropCard.addEventListener('click', () => fileInput.click());

    ['dragenter', 'dragover'].forEach(evtName => {
      window.addEventListener(evtName, (e) => {
        e.preventDefault();
        dropCard.classList.add('dragover');
      });
    });
    ['dragleave', 'drop'].forEach(evtName => {
      window.addEventListener(evtName, (e) => {
        e.preventDefault();
        dropCard.classList.remove('dragover');
      });
    });
    window.addEventListener('drop', async (e) => {
      const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) await openDocxFile(f);
    });

    // Selection & Universal Click-to-Follow inside docPane
    const docPaneEl = document.getElementById('docPane');
    docPaneEl.addEventListener('pointerenter', () => { pointerInDocPane = true; });
    docPaneEl.addEventListener('pointerleave', () => { pointerInDocPane = false; });

    docPaneEl.addEventListener('pointerdown', (e) => {
      if (e.button !== 0) return;
      pointerDownInfo = { x: e.clientX, y: e.clientY, target: e.target };
    });

    docPaneEl.addEventListener('pointerup', (e) => {
      if (e.button !== 0 || !pointerDownInfo || !docState) return;
      const dist = Math.hypot(e.clientX - pointerDownInfo.x, e.clientY - pointerDownInfo.y);
      pointerDownInfo = null;

      if (dist > 6) {
        setTimeout(checkTextSelection, 15);
        return;
      }
      // Genuine click on left content (ignore inline cards when in inline mode, and ignore rev-chip)
      if (e.target.closest('.inline-thread-tray') || e.target.closest('.rev-chip')) return;

      // Clear any accidental 1-char micro-selection from tiny pointer jitter
      const sel = window.getSelection();
      if (sel && !sel.isCollapsed && sel.toString().length <= 2) {
        sel.removeAllRanges();
      } else if (sel && !sel.isCollapsed && sel.toString().length > 2) {
        setTimeout(checkTextSelection, 15);
        return;
      }
      document.getElementById('selToolbar').hidden = true;

      const resolvedId = resolveClickedThreadInDocPane(e.target);
      if (resolvedId) {
        focusThread(resolvedId, 'text', e.clientY);
      }
    });

    // Smooth left-pane scroll auto-follow
    docPaneEl.addEventListener('scroll', () => {
      if (!ui.scrollSync || ui.viewMode !== 'split' || !docState || !pointerInDocPane) return;
      if (performance.now() < suppressDocScrollUntil) return;
      if (scrollSyncTimer) clearTimeout(scrollSyncTimer);
      scrollSyncTimer = setTimeout(syncRightCommentsToLeftScroll, 45);
    }, { passive: true });

    document.addEventListener('keyup', (e) => {
      if (e.shiftKey) setTimeout(checkTextSelection, 15);
    });
    document.addEventListener('mousedown', (e) => {
      const tb = document.getElementById('selToolbar');
      if (!tb.hidden && !tb.contains(e.target)) {
        tb.hidden = true;
      }
    });

    // Keyboard shortcuts: Ctrl+Z / Ctrl+Y, J / K next/prev comment
    document.addEventListener('keydown', (ev) => {
      const tag = (ev.target && ev.target.tagName) || '';
      const inInput = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || ev.target.isContentEditable;
      if ((ev.ctrlKey || ev.metaKey) && !inInput) {
        if (ev.key.toLowerCase() === 'z' && !ev.shiftKey) {
          ev.preventDefault();
          undo();
          return;
        }
        if (ev.key.toLowerCase() === 'y' || (ev.key.toLowerCase() === 'z' && ev.shiftKey)) {
          ev.preventDefault();
          redo();
          return;
        }
      }
      if (inInput || !docState || document.querySelector('dialog[open]')) return;
      if (ev.key === 'j' || ev.key === 'J' || ev.key === 'ArrowDown') {
        const list = getVisibleThreads();
        if (!list.length) return;
        ev.preventDefault();
        const idx = Math.max(0, list.findIndex(t => t.id === ui.activeThreadId));
        const next = list[Math.min(list.length - 1, idx + 1)];
        focusThread(next.id, 'keyboard');
      } else if (ev.key === 'k' || ev.key === 'K' || ev.key === 'ArrowUp') {
        const list = getVisibleThreads();
        if (!list.length) return;
        ev.preventDefault();
        const idx = Math.max(0, list.findIndex(t => t.id === ui.activeThreadId));
        const prev = list[Math.max(0, idx - 1)];
        focusThread(prev.id, 'keyboard');
      }
    });

    setupDialogLightDismissFallbacks();
  }

  /* modern-web-guidance: fallback for <dialog closedby="any"> + returnValue reset */
  function setupDialogLightDismissFallbacks() {
    const dialogs = document.querySelectorAll('dialog');
    dialogs.forEach(dlg => {
      dlg.addEventListener('close', () => {
        dlg.returnValue = '';
      });
    });

    if (!('closedBy' in HTMLDialogElement.prototype)) {
      const outside = (e, dialog) => {
        const r = dialog.getBoundingClientRect();
        return e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom;
      };
      for (const dialog of document.querySelectorAll('dialog[closedby="any"]')) {
        let pressedOutside = false;
        dialog.addEventListener('pointerdown', (e) => { pressedOutside = outside(e, dialog); });
        dialog.addEventListener('pointerup', (e) => {
          if (pressedOutside && outside(e, dialog)) dialog.close();
        });
      }
    }
  }

  /* modern-web-guidance: persistent toast notifications using popover="manual" */
  function notifyToast(msg) {
    // Suppress background toasts while a modal dialog is open
    if (document.querySelector('dialog[open]')) return;

    const liveFallback = document.getElementById('srLiveRegion');
    if (typeof document.body.ariaNotify === 'function') {
      document.body.ariaNotify(msg);
    } else if (liveFallback) {
      liveFallback.textContent = msg;
    }

    const region = document.getElementById('toastRegion');
    const item = document.createElement('div');
    item.className = 'toast-item';
    item.setAttribute('popover', 'manual');

    const span = document.createElement('span');
    span.textContent = msg;

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'tiny-btn';
    closeBtn.textContent = '关闭';
    closeBtn.addEventListener('click', () => {
      if (typeof item.hidePopover === 'function') item.hidePopover();
      item.remove();
    });

    item.appendChild(span);
    item.appendChild(closeBtn);
    region.appendChild(item);
    if (typeof item.showPopover === 'function') {
      try { item.showPopover(); } catch (e) {}
    }

    setTimeout(() => {
      if (!item.isConnected) return;
      if (item.contains(document.activeElement)) return;
      if (typeof item.hidePopover === 'function') {
        try { item.hidePopover(); } catch (e) {}
      }
      item.remove();
    }, 3600);
  }

  /* ==================== FILE OPEN & EXPORT ==================== */
  async function openDocxFile(file) {
    try {
      const buf = await file.arrayBuffer();
      docState = await Engine.loadDocxBuffer(buf, file.name);
      undoStack = [];
      redoStack = [];
      ui.statusFilter = 'ALL';
      ui.authorFilter = 'ALL';
      ui.searchQuery = '';
      document.getElementById('searchInput').value = '';
      ui.activeThreadId = docState.threads.length ? docState.threads[0].id : null;

      document.getElementById('dropStage').hidden = true;
      document.getElementById('subBar').hidden = false;
      document.getElementById('workspace').hidden = false;
      document.getElementById('exportBtn').disabled = false;
      refreshAuthorFilterOptions();
      updateUndoRedoButtons();
      renderWorkspace();
      notifyToast(`已导入 ${file.name}（共 ${docState.threads.length} 条批注线程）`);
    } catch (err) {
      alert('导入 Word 失败：' + (err && err.message ? err.message : String(err)));
    }
  }

  async function exportCurrentDocx() {
    if (!docState) return;
    try {
      const bytes = await Engine.exportDocxBytes(docState);
      const blob = new Blob([bytes], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
      const base = docState.fileName.replace(/\.docx$/i, '');
      const outName = `${base}_已审阅.docx`;
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = outName;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 5000);
      notifyToast(`已导出标准 Word 文档：${outName}`);
    } catch (err) {
      alert('导出 Word 失败：' + (err && err.message ? err.message : String(err)));
    }
  }

  /* ==================== UNDO / REDO ==================== */
  function pushUndoSnapshot() {
    if (!docState) return;
    undoStack.push(Engine.snapshotMutableState(docState));
    if (undoStack.length > 35) undoStack.shift();
    redoStack = [];
    updateUndoRedoButtons();
  }

  function undo() {
    if (!docState || !undoStack.length) return;
    redoStack.push(Engine.snapshotMutableState(docState));
    const snap = undoStack.pop();
    Engine.restoreMutableState(docState, snap);
    refreshAuthorFilterOptions();
    updateUndoRedoButtons();
    renderWorkspace();
    notifyToast('已撤销上一步操作');
  }

  function redo() {
    if (!docState || !redoStack.length) return;
    undoStack.push(Engine.snapshotMutableState(docState));
    const snap = redoStack.pop();
    Engine.restoreMutableState(docState, snap);
    refreshAuthorFilterOptions();
    updateUndoRedoButtons();
    renderWorkspace();
    notifyToast('已重做操作');
  }

  function updateUndoRedoButtons() {
    document.getElementById('undoBtn').disabled = undoStack.length === 0;
    document.getElementById('redoBtn').disabled = redoStack.length === 0;
  }

  /* ==================== FILTERS & MODES ==================== */
  function syncModeButtons() {
    document.body.classList.toggle('view-inline', ui.viewMode === 'inline');
    document.getElementById('btnViewSplit').classList.toggle('active', ui.viewMode === 'split');
    document.getElementById('btnViewInline').classList.toggle('active', ui.viewMode === 'inline');
    document.getElementById('btnEditTrack').classList.toggle('active', ui.editMode === 'track');
    document.getElementById('btnEditDirect').classList.toggle('active', ui.editMode === 'direct');
    const syncBtn = document.getElementById('scrollSyncBtn');
    if (syncBtn) {
      syncBtn.classList.toggle('active', ui.scrollSync);
      syncBtn.textContent = ui.scrollSync ? '🔗 左滚右随：开' : '🔗 左滚右随：关';
    }
  }

  window.toggleScrollSync = function () {
    ui.scrollSync = !ui.scrollSync;
    localStorage.setItem('agy_scroll_sync', ui.scrollSync ? '1' : '0');
    syncModeButtons();
    notifyToast(ui.scrollSync ? '已开启左侧滚动右侧批注自动跟随' : '已关闭左侧滚屏跟随（点击文字仍自动对齐）');
  };

  window.setViewMode = function (m) {
    ui.viewMode = m;
    localStorage.setItem('agy_view_mode', m);
    syncModeButtons();
    renderWorkspace();
  };

  window.setEditMode = function (m) {
    ui.editMode = m;
    localStorage.setItem('agy_edit_mode', m);
    syncModeButtons();
    notifyToast(m === 'track' ? '已切换为：修订模式（改词保留红绿痕迹）' : '已切换为：直接改字模式（无痕修改）');
  };

  window.setStatusFilter = function (f) {
    ui.statusFilter = f;
    document.querySelectorAll('[data-status-filter]').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.statusFilter === f);
    });
    renderWorkspace();
  };

  window.onAuthorFilterChange = function (val) {
    ui.authorFilter = val;
    renderWorkspace();
  };

  window.onSearchInput = function (val) {
    ui.searchQuery = (val || '').trim().toLowerCase();
    renderWorkspace();
  };

  function refreshAuthorFilterOptions() {
    if (!docState) return;
    const authors = new Set();
    for (const th of docState.threads) {
      if (th.author) authors.add(th.author);
      for (const rp of th.replies) {
        if (rp.author) authors.add(rp.author);
      }
    }
    const sel = document.getElementById('authorSelect');
    const prev = ui.authorFilter;
    sel.innerHTML = '<option value="ALL">全部审校/回复人</option>';
    for (const au of authors) {
      const opt = document.createElement('option');
      opt.value = au;
      opt.textContent = `署名：${au}`;
      sel.appendChild(opt);
    }
    sel.value = authors.has(prev) ? prev : 'ALL';
    ui.authorFilter = sel.value;
  }

  function getThreadAnchorText(th) {
    const cell = docState.rows[th.anchor.row]?.cells[th.anchor.col];
    const p = cell?.paras[th.anchor.pIdx];
    if (!p) return '';
    return Engine.paraVisibleText(p).slice(th.anchor.start, th.anchor.end);
  }

  function getThreadLinkedEnPreview(th) {
    const parts = [];
    for (const lk of th.linkedSpans) {
      const p = docState.rows[lk.row]?.cells[lk.col]?.paras[lk.pIdx];
      if (p) {
        const s = Engine.paraVisibleText(p).slice(lk.start, lk.end).trim();
        if (s) parts.push(s);
      }
    }
    return parts.join(' / ');
  }

  function isThreadVisible(th) {
    if (ui.statusFilter === 'OPEN' && th.done) return false;
    if (ui.statusFilter === 'DONE' && !th.done) return false;
    if (ui.statusFilter === 'REPLIED' && th.replies.length === 0) return false;

    if (ui.authorFilter !== 'ALL') {
      const matchRoot = th.author === ui.authorFilter;
      const matchReply = th.replies.some(r => r.author === ui.authorFilter);
      if (!matchRoot && !matchReply) return false;
    }

    if (ui.searchQuery) {
      const hay = [
        getThreadAnchorText(th),
        getThreadLinkedEnPreview(th),
        th.text,
        th.author,
        ...th.replies.map(r => r.text + ' ' + r.author)
      ].join(' ').toLowerCase();
      if (!hay.includes(ui.searchQuery)) return false;
    }
    return true;
  }

  function getVisibleThreads() {
    return docState ? docState.threads.filter(isThreadVisible) : [];
  }

  /* ==================== RENDERING ==================== */
  function escHtml(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function locateSentenceBlockForThread(th) {
    const rObj = docState.rows[th.anchor.row];
    if (!rObj) return 0;
    const cOff = Engine.paraOffsetToCellOffset(rObj.cells[th.anchor.col], th.anchor.pIdx, th.anchor.start);
    const idx = rObj.blocks.findIndex(b => {
      const sp = th.anchor.col === 0 ? b.zhSpan : b.enSpan;
      return cOff >= sp[0] && cOff < sp[1];
    });
    return idx >= 0 ? idx : 0;
  }

  /**
   * Render exact slice [cellStart, cellEnd) of (rowIdx, colIdx) preserving token highlights,
   * Track Changes, comment pins, and strong active spotlight zoom!
   */
  function renderCellSliceHtml(rowIdx, colIdx, cellStart, cellEnd, activeThread, threadOrderMap) {
    const cell = docState.rows[rowIdx].cells[colIdx];
    if (!cell) return '';

    let cellCursor = 0;
    let htmlOut = '';

    cell.paras.forEach((pModel, pIdx) => {
      const pText = Engine.paraVisibleText(pModel);
      const pLen = pText.length;
      const pCellStart = cellCursor;
      const pCellEnd = cellCursor + pLen;
      cellCursor = pCellEnd + 1;

      if (pCellEnd < cellStart || pCellStart >= cellEnd) {
        if (!(pLen === 0 && pCellStart >= cellStart && pCellStart <= cellEnd)) return;
      }

      const localSliceStart = Math.max(0, cellStart - pCellStart);
      const localSliceEnd = Math.min(pLen, cellEnd - pCellStart);

      // Collect cut points inside [localSliceStart, localSliceEnd]
      const cuts = new Set([localSliceStart, localSliceEnd]);

      // Token boundaries
      let tPos = 0;
      for (const tok of pModel.tokens) {
        const len = Engine.tokenLen(tok);
        if (tPos > localSliceStart && tPos < localSliceEnd) cuts.add(tPos);
        if (tPos + len > localSliceStart && tPos + len < localSliceEnd) cuts.add(tPos + len);
        tPos += len;
      }

      // Thread anchors & active spotlight spans in this paragraph
      const paraThreads = docState.threads.filter(
        th => th.anchor.row === rowIdx && th.anchor.col === colIdx && th.anchor.pIdx === pIdx
      );
      const pinsAtOffset = new Map();
      for (const th of paraThreads) {
        if (th.anchor.start >= localSliceStart && th.anchor.start <= localSliceEnd) cuts.add(th.anchor.start);
        if (th.anchor.end >= localSliceStart && th.anchor.end <= localSliceEnd) {
          cuts.add(th.anchor.end);
          if (!pinsAtOffset.has(th.anchor.end)) pinsAtOffset.set(th.anchor.end, []);
          pinsAtOffset.get(th.anchor.end).push(th);
        }
      }

      const activeFocusSpans = [];
      if (activeThread) {
        if (activeThread.anchor.row === rowIdx && activeThread.anchor.col === colIdx && activeThread.anchor.pIdx === pIdx) {
          activeFocusSpans.push({ start: activeThread.anchor.start, end: activeThread.anchor.end, kind: 'cyan' });
        }
        for (const lk of activeThread.linkedSpans) {
          if (lk.row === rowIdx && lk.col === colIdx && lk.pIdx === pIdx) {
            activeFocusSpans.push({
              start: lk.start,
              end: lk.end,
              kind: lk.color === 'green' ? 'green' : 'cyan'
            });
          }
        }
        for (const sp of activeFocusSpans) {
          const a = Math.max(localSliceStart, sp.start);
          const b = Math.min(localSliceEnd, sp.end);
          if (a < b) { cuts.add(a); cuts.add(b); }
        }
      }

      // Map any character in this paragraph to its clickable thread (either anchor or linkedSpan)
      const findThreadForOffset = (off) => {
        const direct = paraThreads.find(th => off >= th.anchor.start && off < th.anchor.end);
        if (direct) return direct.id;
        const linked = docState.threads.find(th =>
          th.linkedSpans.some(lk => lk.row === rowIdx && lk.col === colIdx && lk.pIdx === pIdx && off >= lk.start && off < lk.end)
        );
        return linked ? linked.id : '';
      };

      const renderPinsAt = (off) => {
        const list = pinsAtOffset.get(off);
        if (!list) return '';
        return list.map(th => {
          const ord = threadOrderMap.get(th.id) || '?';
          const actCls = (activeThread && activeThread.id === th.id) ? 'pin-active' : '';
          const doneCls = th.done ? 'pin-done' : '';
          const repBadge = th.replies.length ? `·${th.replies.length + 1}` : '';
          return `<sup class="cmt-pin ${actCls} ${doneCls}" onclick="window.onPinOrSpanClick(event, '${th.id}')" title="点击查看并放大第 #${ord} 号批注">${ord}${repBadge}</sup>`;
        }).join('');
      };

      const sortedCuts = Array.from(cuts).sort((a, b) => a - b);

      // Walk tokens and sub-slices cleanly
      let tokCursor = 0;
      let tokIdx = 0;

      if (pinsAtOffset.has(localSliceStart)) {
        // Only emit zero-length anchor pins at start
        const zeroThreads = pinsAtOffset.get(localSliceStart).filter(th => th.anchor.start === th.anchor.end);
        if (zeroThreads.length) {
          htmlOut += zeroThreads.map(th => {
            const ord = threadOrderMap.get(th.id) || '?';
            return `<sup class="cmt-pin" data-thread-id="${th.id}">${ord}</sup>`;
          }).join('');
        }
      }

      for (let i = 0; i < sortedCuts.length - 1; i++) {
        const segA = sortedCuts[i];
        const segB = sortedCuts[i + 1];
        if (segA >= segB) continue;

        while (tokIdx < pModel.tokens.length && tokCursor + Engine.tokenLen(pModel.tokens[tokIdx]) <= segA) {
          tokCursor += Engine.tokenLen(pModel.tokens[tokIdx]);
          tokIdx++;
        }
        const tok = pModel.tokens[tokIdx];
        if (!tok) continue;

        const focusHit = activeFocusSpans.find(sp => sp.start <= segA && segB <= sp.end);
        const linkedThreadId = findThreadForOffset(segA);

        const cls = ['doc-seg'];
        if (tok.hl === 'cyan') cls.push('hl-cyan');
        else if (tok.hl === 'green') cls.push('hl-green');
        else if (tok.hl === 'yellow') cls.push('hl-yellow');

        if (focusHit) {
          cls.push(focusHit.kind === 'green' ? 'hl-focus-green' : 'hl-focus-cyan');
        }

        const thAttr = linkedThreadId ? `data-thread-id="${linkedThreadId}"` : '';

        if (tok.kind === 'rev') {
          htmlOut += `<span class="${cls.join(' ')} rev-chip" data-row="${rowIdx}" data-col="${colIdx}" data-p="${pIdx}" data-start="${segA}" data-end="${segB}" onclick="window.onRevTokenClick(event, ${rowIdx}, ${colIdx}, ${pIdx}, '${tok.id}')" title="修订人：${escHtml(tok.author)}（点击可接受/拒绝此修订）">`;
          if (tok.delText) htmlOut += `<span class="rev-del">${escHtml(tok.delText)}</span>`;
          if (tok.insText) htmlOut += `<span class="rev-ins">${escHtml(tok.insText)}</span>`;
          htmlOut += `</span>`;
        } else {
          const subStr = tok.text.slice(segA - tokCursor, segB - tokCursor);
          htmlOut += `<span class="${cls.join(' ')}" data-row="${rowIdx}" data-col="${colIdx}" data-p="${pIdx}" data-start="${segA}" data-end="${segB}" ${thAttr}>${escHtml(subStr)}</span>`;
        }

        // Emit pins whose anchor.end === segB (excluding zero-length already emitted)
        if (pinsAtOffset.has(segB)) {
          const nonZero = pinsAtOffset.get(segB).filter(th => th.anchor.end > th.anchor.start);
          if (nonZero.length) {
            htmlOut += nonZero.map(th => {
              const ord = threadOrderMap.get(th.id) || '?';
              const actCls = (activeThread && activeThread.id === th.id) ? 'pin-active' : '';
              const doneCls = th.done ? 'pin-done' : '';
              const repBadge = th.replies.length ? `·${th.replies.length + 1}` : '';
              return `<sup class="cmt-pin ${actCls} ${doneCls}" data-thread-id="${th.id}" title="点击定位并放大第 #${ord} 号批注">${ord}${repBadge}</sup>`;
            }).join('');
          }
        }
      }

      if (pCellEnd < cellEnd) htmlOut += '<br>';
    });

    return htmlOut;
  }

  function formatShortDate(iso) {
    if (!iso) return '';
    return iso.replace('T', ' ').replace(/:\d{2}Z$/, '').slice(0, 16);
  }

  function buildThreadCardHtml(th, ordNum, paneKind) {
    const pk = paneKind || ui.viewMode;
    const isAct = ui.activeThreadId === th.id;
    const rObj = docState.rows[th.anchor.row];
    const bIdx = locateSentenceBlockForThread(th);
    const blkLabel = rObj?.blocks[bIdx]?.label || 'S1';
    const locStr = docState.rows.length > 1 ? `第${th.anchor.row + 1}段·${blkLabel}` : blkLabel;
    const anchorTxt = getThreadAnchorText(th) || '（定点批注）';
    const enPreview = getThreadLinkedEnPreview(th);

    const repliesHtml = th.replies.length
      ? `<div class="replies-list">` + th.replies.map(rp => `
          <div class="reply-bubble">
            <div class="reply-head">
              <span><span class="reply-author">${escHtml(rp.author)}</span> · ${escHtml(formatShortDate(rp.date))}</span>
              <span>
                <button type="button" class="tiny-btn" onclick="window.openEditReplyModal(event, '${th.id}', '${rp.id}')">改</button>
                <button type="button" class="tiny-btn danger" onclick="window.deleteReply(event, '${th.id}', '${rp.id}')">删</button>
              </span>
            </div>
            <div style="white-space:pre-wrap">${escHtml(rp.text)}</div>
          </div>
        `).join('') + `</div>`
      : '';

    return `
      <article class="thread-card ${isAct ? 'active' : ''} ${th.done ? 'thread-done' : ''}" id="card-${pk}-${th.id}" onclick="window.focusThread('${th.id}', 'card')">
        <div class="card-top">
          <div class="card-badges">
            <span class="num-pill">#${ordNum}</span>
            <span class="author-tag">${escHtml(th.author)}</span>
            <span class="loc-tag">${escHtml(locStr)}${th.done ? ' · 已解决' : ''}</span>
          </div>
          <div onclick="event.stopPropagation()" style="display:flex;gap:4px">
            <button type="button" class="tiny-btn" onclick="window.toggleThreadDone('${th.id}')">${th.done ? '重开' : '✓ 解决'}</button>
            <button type="button" class="tiny-btn" onclick="window.openEditCommentModal('${th.id}')">改字</button>
            <button type="button" class="tiny-btn danger" onclick="window.openDeleteCommentModal('${th.id}')">删除</button>
          </div>
        </div>
        <div class="anchor-strip">
          <span class="anchor-zh">【${escHtml(anchorTxt.slice(0, 34))}】</span>
          ${enPreview ? `<span class="anchor-en">${escHtml(enPreview.slice(0, 48))}</span>` : ''}
        </div>
        <div class="cmt-text">${escHtml(th.text)}</div>
        ${repliesHtml}
        <div class="reply-composer" onclick="event.stopPropagation()">
          <input type="text" class="reply-input" id="reply-in-${pk}-${th.id}" placeholder="添加多轮回复（按 Enter 发送）…" onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();window.submitReply('${th.id}');}">
          <button type="button" class="tiny-btn" onclick="window.submitReply('${th.id}')">回复</button>
        </div>
      </article>
    `;
  }

  function renderWorkspace() {
    if (!docState) return;
    Engine.sortThreadsInPlace(docState.threads, docState.rows);

    const threadOrderMap = new Map();
    docState.threads.forEach((th, i) => threadOrderMap.set(th.id, i + 1));

    const visibleThreads = getVisibleThreads();
    const activeThread = docState.threads.find(t => t.id === ui.activeThreadId) || null;
    const activeBlockIdx = activeThread ? locateSentenceBlockForThread(activeThread) : -1;

    // Header badge
    const replyTotal = docState.threads.reduce((s, t) => s + t.replies.length, 0);
    document.getElementById('docBadge').textContent =
      `${docState.fileName} · ${docState.threads.length} 条批注` + (replyTotal ? ` · ${replyTotal} 条回复` : '');
    document.getElementById('filterCountLabel').textContent =
      `当前显示 ${visibleThreads.length} / ${docState.threads.length} 条`;

    // Render left document pane
    let docHtml = '';
    docState.rows.forEach((rObj, rIdx) => {
      if (docState.rows.length > 1) {
        docHtml += `<h2 class="row-header">第 ${rIdx + 1} 段</h2>`;
      }
      const hasCol1 = !!rObj.cells[1];
      rObj.blocks.forEach((blk, bIdx) => {
        const isFocused = activeThread && activeThread.anchor.row === rIdx && activeBlockIdx === bIdx;
        const leftHtml = renderCellSliceHtml(rIdx, 0, blk.zhSpan[0], blk.zhSpan[1], activeThread, threadOrderMap);
        const rightHtml = hasCol1
          ? renderCellSliceHtml(rIdx, 1, blk.enSpan[0], blk.enSpan[1], activeThread, threadOrderMap)
          : '';

        const blkThreads = ui.viewMode === 'inline'
          ? visibleThreads.filter(th => th.anchor.row === rIdx && locateSentenceBlockForThread(th) === bIdx)
          : [];
        const trayHtml = blkThreads.map(th => buildThreadCardHtml(th, threadOrderMap.get(th.id), 'inline')).join('');

        docHtml += `
          <div class="sent-row ${isFocused ? 'active-sent' : ''}" id="sblk-${rIdx}-${bIdx}" data-row="${rIdx}" data-bidx="${bIdx}">
            <div class="sent-grid ${hasCol1 ? '' : 'single-col'}">
              <div class="sent-cell"><span class="sent-idx">${escHtml(blk.label)}</span>${leftHtml}</div>
              ${hasCol1 ? `<div class="sent-cell">${rightHtml}</div>` : ''}
            </div>
            <div class="inline-thread-tray ${blkThreads.length ? 'has-items' : ''}">${trayHtml}</div>
          </div>
        `;
      });
    });

    const dPane = document.getElementById('docPane');
    const cPane = document.getElementById('commentsPane');
    const prevDocScroll = dPane.scrollTop;
    const prevCmtScroll = cPane.scrollTop;

    dPane.innerHTML = docHtml;
    dPane.scrollTop = prevDocScroll;

    // Render right comments pane without losing scroll offset
    cPane.classList.toggle('has-active', !!activeThread);
    if (ui.viewMode === 'split') {
      cPane.innerHTML = visibleThreads.length
        ? visibleThreads.map(th => buildThreadCardHtml(th, threadOrderMap.get(th.id), 'split')).join('')
        : `<div style="padding:36px 20px;text-align:center;color:var(--muted)">当前筛选条件下暂无批注（可在左侧划选文字新增批注）</div>`;
      cPane.scrollTop = prevCmtScroll;
    }
  }

  /* ==================== UNIVERSAL LEFT-CLICK & SCROLL ALIGNMENT ==================== */
  function spanDist(pos, a, b) {
    if (pos >= a && pos <= b) return 0;
    return pos < a ? (a - pos) : (pos - b);
  }

  function resolveClickedThreadInDocPane(targetEl) {
    if (!docState) return null;
    const vis = getVisibleThreads();
    if (!vis.length) return null;

    // 1. Direct hit on a pin, anchor highlight, or linked English highlight
    const directEl = targetEl.closest && targetEl.closest('[data-thread-id]');
    if (directEl && directEl.dataset.threadId) {
      const id = directEl.dataset.threadId;
      if (vis.some(t => t.id === id)) return id;
    }

    // 2. Click on ANY plain Chinese or English text inside a sentence block -> pick nearest comment!
    const segEl = targetEl.closest && targetEl.closest('.doc-seg');
    const sentEl = targetEl.closest && targetEl.closest('.sent-row');
    if (!sentEl) return null;

    const rIdx = Number(sentEl.dataset.row);
    const bIdx = Number(sentEl.dataset.bidx);
    const inBlock = vis.filter(th => th.anchor.row === rIdx && locateSentenceBlockForThread(th) === bIdx);

    if (inBlock.length === 1) return inBlock[0].id;

    if (inBlock.length > 1 && segEl) {
      const col = Number(segEl.dataset.col);
      const pIdx = Number(segEl.dataset.p);
      const midOff = (Number(segEl.dataset.start) + Number(segEl.dataset.end)) / 2;

      let bestTh = inBlock[0];
      let bestDist = Infinity;
      for (const th of inBlock) {
        let d = Infinity;
        if (col === th.anchor.col && pIdx === th.anchor.pIdx) {
          d = spanDist(midOff, th.anchor.start, th.anchor.end);
        } else {
          for (const lk of th.linkedSpans) {
            if (lk.row === rIdx && lk.col === col && lk.pIdx === pIdx) {
              d = Math.min(d, spanDist(midOff, lk.start, lk.end));
            }
          }
        }
        if (d < bestDist) {
          bestDist = d;
          bestTh = th;
        }
      }
      return bestTh.id;
    }

    if (inBlock.length > 0) return inBlock[0].id;

    // 3. Fallback when clicked sentence itself has 0 comments: pick nearest comment by (row, blockIdx)
    let nearest = vis[0];
    let minGap = Infinity;
    for (const th of vis) {
      const thBlk = locateSentenceBlockForThread(th);
      const gap = Math.abs((th.anchor.row - rIdx) * 100 + (thBlk - bIdx));
      if (gap < minGap) {
        minGap = gap;
        nearest = th;
      }
    }
    return nearest.id;
  }

  function alignCardInRightPane(threadId, targetClientY, durationMs) {
    const cPane = document.getElementById('commentsPane');
    const cardEl = document.getElementById(`card-split-${threadId}`);
    if (!cPane || !cardEl) return;

    const cRect = cPane.getBoundingClientRect();
    const cardRect = cardEl.getBoundingClientRect();

    // Determine desired vertical position of card relative to commentsPane top
    let desiredVisualTop = 100;
    if (typeof targetClientY === 'number') {
      desiredVisualTop = targetClientY - cRect.top - 26;
    } else {
      const focusSpan = document.querySelector('#docPane .hl-focus-cyan');
      if (focusSpan) {
        desiredVisualTop = focusSpan.getBoundingClientRect().top - cRect.top - 18;
      }
    }
    desiredVisualTop = Math.max(18, Math.min(Math.max(60, cPane.clientHeight - 180), desiredVisualTop));

    const targetScrollTop = cPane.scrollTop + (cardRect.top - cRect.top) - desiredVisualTop;
    smoothScrollElementTo(cPane, targetScrollTop, durationMs || 260);
  }

  function syncRightCommentsToLeftScroll() {
    if (!docState || ui.viewMode !== 'split') return;
    const vis = getVisibleThreads();
    if (!vis.length) return;
    const dPane = document.getElementById('docPane');
    const dRect = dPane.getBoundingClientRect();
    const probeY = dRect.top + Math.min(140, dPane.clientHeight * 0.26);

    const rowsEls = dPane.querySelectorAll('.sent-row');
    let hitRow = null;
    for (const el of rowsEls) {
      const r = el.getBoundingClientRect();
      if (r.bottom >= probeY) {
        hitRow = el;
        break;
      }
    }
    if (!hitRow) return;

    const rIdx = Number(hitRow.dataset.row);
    const bIdx = Number(hitRow.dataset.bidx);
    const match = vis.find(th => th.anchor.row === rIdx && locateSentenceBlockForThread(th) === bIdx);
    if (!match || match.id === ui.activeThreadId) return;

    ui.activeThreadId = match.id;
    renderWorkspace();
    alignCardInRightPane(match.id, hitRow.getBoundingClientRect().top + 20, 210);
  }

  window.focusThread = function (threadId, source, clickedClientY) {
    ui.activeThreadId = threadId;
    renderWorkspace();
    const th = docState.threads.find(t => t.id === threadId);
    if (!th) return;

    if (ui.viewMode === 'split') {
      if (source === 'text') {
        suppressDocScrollUntil = performance.now() + 550;
        alignCardInRightPane(threadId, clickedClientY, 260);
      } else if (source === 'card') {
        suppressDocScrollUntil = performance.now() + 450;
        const dPane = document.getElementById('docPane');
        const cardEl = document.getElementById(`card-split-${threadId}`);
        const focusSpan = dPane.querySelector('.hl-focus-cyan') || document.getElementById(`sblk-${th.anchor.row}-${locateSentenceBlockForThread(th)}`);
        if (dPane && focusSpan) {
          const dRect = dPane.getBoundingClientRect();
          const fRect = focusSpan.getBoundingClientRect();
          const cardVisualY = cardEl ? Math.max(40, Math.min(dPane.clientHeight - 150, cardEl.getBoundingClientRect().top - dRect.top)) : 120;
          const targetDocTop = dPane.scrollTop + (fRect.top - dRect.top) - cardVisualY;
          smoothScrollElementTo(dPane, targetDocTop, 260);
        }
      } else if (source === 'keyboard') {
        suppressDocScrollUntil = performance.now() + 450;
        const dPane = document.getElementById('docPane');
        const focusSpan = dPane.querySelector('.hl-focus-cyan') || document.getElementById(`sblk-${th.anchor.row}-${locateSentenceBlockForThread(th)}`);
        if (dPane && focusSpan) {
          const dRect = dPane.getBoundingClientRect();
          const fRect = focusSpan.getBoundingClientRect();
          const targetDocTop = dPane.scrollTop + (fRect.top - dRect.top) - Math.round(dPane.clientHeight * 0.28);
          smoothScrollElementTo(dPane, targetDocTop, 220);
        }
        requestAnimationFrame(() => alignCardInRightPane(threadId, undefined, 220));
      }
    }
  };

  /* ==================== TEXT SELECTION & FLOATING TOOLBAR ==================== */
  function resolveBoundaryPoint(node, domOffset) {
    let cur = node;
    while (cur && cur !== document.body) {
      if (cur.nodeType === 1 && cur.classList && cur.classList.contains('doc-seg')) {
        const row = Number(cur.dataset.row);
        const col = Number(cur.dataset.col);
        const pIdx = Number(cur.dataset.p);
        const segStart = Number(cur.dataset.start);
        const segEnd = Number(cur.dataset.end);
        if (cur.classList.contains('rev-chip')) {
          return { row, col, pIdx, off: domOffset > 0 ? segEnd : segStart };
        }
        return { row, col, pIdx, off: Math.min(segEnd, segStart + domOffset) };
      }
      cur = cur.parentNode;
    }
    return null;
  }

  function checkTextSelection() {
    const tb = document.getElementById('selToolbar');
    if (!docState) { tb.hidden = true; return; }
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || !sel.rangeCount) {
      tb.hidden = true;
      return;
    }
    const range = sel.getRangeAt(0);
    const docPane = document.getElementById('docPane');
    if (!docPane.contains(range.commonAncestorContainer)) {
      tb.hidden = true;
      return;
    }

    const ptA = resolveBoundaryPoint(range.startContainer, range.startOffset);
    const ptB = resolveBoundaryPoint(range.endContainer, range.endOffset);
    if (!ptA || !ptB || ptA.row !== ptB.row || ptA.col !== ptB.col || ptA.pIdx !== ptB.pIdx) {
      tb.hidden = true;
      return;
    }

    const start = Math.min(ptA.off, ptB.off);
    const end = Math.max(ptA.off, ptB.off);
    if (end <= start) { tb.hidden = true; return; }

    const pModel = docState.rows[ptA.row].cells[ptA.col].paras[ptA.pIdx];
    const selectedText = Engine.paraVisibleText(pModel).slice(start, end);
    const rect = range.getBoundingClientRect();

    ui.currentSelection = { row: ptA.row, col: ptA.col, pIdx: ptA.pIdx, start, end, text: selectedText };

    // Update mode-aware button label & bind button visibility
    document.getElementById('selEditTextBtn').textContent =
      ui.editMode === 'track' ? '✏️ 修订改词' : '📝 直接改字';
    document.getElementById('selBindCyanBtn').hidden = !ui.activeThreadId;
    document.getElementById('selBindGreenBtn').hidden = !ui.activeThreadId;

    tb.hidden = false;
    const top = Math.max(10, rect.top - tb.offsetHeight - 10);
    const left = Math.max(12, Math.min(window.innerWidth - tb.offsetWidth - 12, rect.left + rect.width / 2 - tb.offsetWidth / 2));
    tb.style.top = `${top}px`;
    tb.style.left = `${left}px`;
  }

  /* ==================== HIGHLIGHT, BIND, ADD/EDIT/DELETE/REPLY ==================== */
  window.applySelectionHighlight = function (color) {
    const s = ui.currentSelection;
    if (!s || !docState) return;
    pushUndoSnapshot();
    Engine.setParaHighlight(docState, s.row, s.col, s.pIdx, s.start, s.end, color);
    // If clearing highlight, also remove matching linkedSpans
    if (!color) {
      for (const th of docState.threads) {
        th.linkedSpans = th.linkedSpans.filter(lk =>
          !(lk.row === s.row && lk.col === s.col && lk.pIdx === s.pIdx && lk.start < s.end && lk.end > s.start)
        );
      }
    }
    window.getSelection()?.removeAllRanges();
    document.getElementById('selToolbar').hidden = true;
    renderWorkspace();
  };

  window.bindSelectionToActiveThread = function (color) {
    const s = ui.currentSelection;
    const th = docState?.threads.find(t => t.id === ui.activeThreadId);
    if (!s || !th) return;
    pushUndoSnapshot();
    Engine.setParaHighlight(docState, s.row, s.col, s.pIdx, s.start, s.end, color);
    th.linkedSpans.push({ row: s.row, col: s.col, pIdx: s.pIdx, start: s.start, end: s.end, color });
    window.getSelection()?.removeAllRanges();
    document.getElementById('selToolbar').hidden = true;
    renderWorkspace();
    notifyToast(`已将选中词绑定到当前批注（${color === 'green' ? '绿色参照' : '青色对应词'}）`);
  };

  window.openNewCommentModalFromSelection = function () {
    const s = ui.currentSelection;
    if (!s || !docState) return;
    document.getElementById('selToolbar').hidden = true;
    ui.pendingModalAction = { type: 'new-comment', sel: { ...s } };
    document.getElementById('editModalTitle').textContent = `新增批注 — 【${s.text.slice(0, 22)}】`;
    const ta = document.getElementById('editModalTextarea');
    ta.value = '';
    const dlg = document.getElementById('editModal');
    dlg.showModal();
    ta.focus();
  };

  window.openEditCommentModal = function (threadId) {
    const th = docState?.threads.find(t => t.id === threadId);
    if (!th) return;
    ui.pendingModalAction = { type: 'edit-comment', threadId };
    document.getElementById('editModalTitle').textContent = '修改批注内容';
    const ta = document.getElementById('editModalTextarea');
    ta.value = th.text;
    const dlg = document.getElementById('editModal');
    dlg.showModal();
    ta.focus();
  };

  window.openEditReplyModal = function (ev, threadId, replyId) {
    ev.stopPropagation();
    const th = docState?.threads.find(t => t.id === threadId);
    const rp = th?.replies.find(r => r.id === replyId);
    if (!rp) return;
    ui.pendingModalAction = { type: 'edit-reply', threadId, replyId };
    document.getElementById('editModalTitle').textContent = `修改 ${rp.author} 的回复`;
    const ta = document.getElementById('editModalTextarea');
    ta.value = rp.text;
    const dlg = document.getElementById('editModal');
    dlg.showModal();
    ta.focus();
  };

  window.openEditTextModalFromSelection = function () {
    const s = ui.currentSelection;
    if (!s || !docState) return;
    document.getElementById('selToolbar').hidden = true;
    ui.pendingModalAction = { type: 'edit-text', sel: { ...s }, mode: ui.editMode };
    document.getElementById('editModalTitle').textContent =
      ui.editMode === 'track'
        ? '修订模式改词（留空即修订删除所选词）'
        : '直接无痕修改文字';
    const ta = document.getElementById('editModalTextarea');
    ta.value = s.text;
    const dlg = document.getElementById('editModal');
    dlg.showModal();
    ta.focus();
    ta.select();
  };

  window.confirmEditModal = function () {
    const act = ui.pendingModalAction;
    const val = document.getElementById('editModalTextarea').value;
    const dlg = document.getElementById('editModal');
    if (!act || !docState) { dlg.close(); return; }

    if (act.type === 'new-comment') {
      const trimmed = val.trim();
      if (!trimmed) { alert('请输入批注内容'); return; }
      pushUndoSnapshot();
      const s = act.sel;
      Engine.setParaHighlight(docState, s.row, s.col, s.pIdx, s.start, s.end, 'cyan');
      const newTh = {
        id: Engine.nextUid('th'),
        cid: String(Date.now()),
        paraId: Engine.genHexId(docState.usedHexIds),
        author: ui.reviewer,
        initials: Engine.makeInitials(ui.reviewer),
        date: Engine.nowIsoUtc(),
        text: trimmed,
        done: false,
        anchor: { row: s.row, col: s.col, pIdx: s.pIdx, start: s.start, end: s.end },
        linkedSpans: [],
        replies: []
      };
      docState.threads.push(newTh);
      Engine.sortThreadsInPlace(docState.threads, docState.rows);
      ui.activeThreadId = newTh.id;
      refreshAuthorFilterOptions();
      dlg.close();
      renderWorkspace();
      focusThread(newTh.id, 'text');
      notifyToast('已新增批注（可继续选中右侧英文词点“绑到当前批注”）');
      return;
    }

    if (act.type === 'edit-comment') {
      const th = docState.threads.find(t => t.id === act.threadId);
      if (th && val.trim()) {
        pushUndoSnapshot();
        th.text = val.trim();
        renderWorkspace();
      }
      dlg.close();
      return;
    }

    if (act.type === 'edit-reply') {
      const th = docState.threads.find(t => t.id === act.threadId);
      const rp = th?.replies.find(r => r.id === act.replyId);
      if (rp && val.trim()) {
        pushUndoSnapshot();
        rp.text = val.trim();
        renderWorkspace();
      }
      dlg.close();
      return;
    }

    if (act.type === 'edit-text') {
      pushUndoSnapshot();
      const s = act.sel;
      Engine.editParagraphRange(docState, s.row, s.col, s.pIdx, s.start, s.end, val, act.mode, ui.reviewer);
      window.getSelection()?.removeAllRanges();
      dlg.close();
      renderWorkspace();
      notifyToast(act.mode === 'track' ? '已生成 Word 修订痕迹' : '已直接修改文字');
    }
  };

  /* Delete Comment Modal with 1-click Highlight Cleanup */
  window.openDeleteCommentModal = function (threadId) {
    const th = docState?.threads.find(t => t.id === threadId);
    if (!th) return;
    ui.pendingModalAction = { type: 'delete-comment', threadId };
    document.getElementById('delAnchorPreview').textContent = getThreadAnchorText(th).slice(0, 30);
    document.getElementById('delClearHlCheck').checked = true;
    document.getElementById('deleteModal').showModal();
  };

  window.confirmDeleteComment = function () {
    const act = ui.pendingModalAction;
    const clearHl = document.getElementById('delClearHlCheck').checked;
    const dlg = document.getElementById('deleteModal');
    if (!act || act.type !== 'delete-comment' || !docState) { dlg.close(); return; }

    const idx = docState.threads.findIndex(t => t.id === act.threadId);
    if (idx < 0) { dlg.close(); return; }
    pushUndoSnapshot();
    const th = docState.threads[idx];

    if (clearHl) {
      Engine.setParaHighlight(docState, th.anchor.row, th.anchor.col, th.anchor.pIdx, th.anchor.start, th.anchor.end, null);
      for (const lk of th.linkedSpans) {
        Engine.setParaHighlight(docState, lk.row, lk.col, lk.pIdx, lk.start, lk.end, null);
      }
    }

    docState.threads.splice(idx, 1);
    if (ui.activeThreadId === th.id) {
      ui.activeThreadId = docState.threads[Math.min(idx, docState.threads.length - 1)]?.id || null;
    }
    refreshAuthorFilterOptions();
    dlg.close();
    renderWorkspace();
    notifyToast(clearHl ? '已删除批注并自动清除关联中英文高亮（Ctrl+Z 可撤销）' : '已删除批注');
  };

  window.toggleThreadDone = function (threadId) {
    const th = docState?.threads.find(t => t.id === threadId);
    if (!th) return;
    pushUndoSnapshot();
    th.done = !th.done;
    renderWorkspace();
  };

  window.submitReply = function (threadId) {
    const th = docState?.threads.find(t => t.id === threadId);
    const input = document.getElementById(`reply-in-${ui.viewMode}-${threadId}`);
    if (!th || !input) return;
    const txt = input.value.trim();
    if (!txt) return;
    pushUndoSnapshot();
    th.replies.push({
      id: Engine.nextUid('rp'),
      cid: String(Date.now()),
      paraId: Engine.genHexId(docState.usedHexIds),
      author: ui.reviewer,
      initials: Engine.makeInitials(ui.reviewer),
      date: Engine.nowIsoUtc(),
      text: txt
    });
    input.value = '';
    refreshAuthorFilterOptions();
    renderWorkspace();
    notifyToast('已添加批注回复');
  };

  window.deleteReply = function (ev, threadId, replyId) {
    ev.stopPropagation();
    const th = docState?.threads.find(t => t.id === threadId);
    if (!th) return;
    pushUndoSnapshot();
    th.replies = th.replies.filter(r => r.id !== replyId);
    refreshAuthorFilterOptions();
    renderWorkspace();
    notifyToast('已删除该条回复');
  };

  /* Track Change Accept / Reject Modal */
  window.onRevTokenClick = function (ev, row, col, pIdx, revId) {
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed && sel.toString().trim().length > 0) return;
    ev.stopPropagation();
    const p = docState.rows[row].cells[col].paras[pIdx];
    const tok = p.tokens.find(t => t.id === revId);
    if (!tok) return;
    ui.pendingModalAction = { type: 'resolve-rev', row, col, pIdx, revId };
    document.getElementById('revAuthorSpan').textContent = `${tok.author} · ${formatShortDate(tok.date)}`;
    document.getElementById('revDelPreview').textContent = tok.delText || '（无删除）';
    document.getElementById('revInsPreview').textContent = tok.insText || '（纯删除）';
    document.getElementById('revModal').showModal();
  };

  window.handleRevDecision = function (accept) {
    const act = ui.pendingModalAction;
    const dlg = document.getElementById('revModal');
    if (!act || act.type !== 'resolve-rev' || !docState) { dlg.close(); return; }
    pushUndoSnapshot();
    Engine.resolveTrackToken(docState, act.row, act.col, act.pIdx, act.revId, accept);
    dlg.close();
    renderWorkspace();
    notifyToast(accept ? '已接受该处修订' : '已拒绝该处修订并恢复原文');
  };

  window.triggerOpenFile = function () {
    document.getElementById('docxFileInput').click();
  };

  window.openHelpModal = function () {
    const dlg = document.getElementById('helpModal');
    if (dlg) dlg.showModal();
  };

  window.loadBuiltInDemo = async function () {
    // Construct a clean minimal valid .docx archive in memory with 2 rows & 12 comments for instant interactive testing
    const wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const w14Ns = 'http://schemas.microsoft.com/office/word/2010/wordml';
    const w15Ns = 'http://schemas.microsoft.com/office/word/2012/wordml';
    const pairs = [
      [
        '我在这里与你们说话并不是普通的谈话，你们需慎重对待才是，更需对我的话多作揣摩。谈到神的性情你们更是不愿多多斟酌一番，很少有人在这个问题上下功夫。难怪你们都目中无神，生活中无真理，这样的一班人怎么能称作是圣徒呢？既然你们对此知之甚少，那我就只好多费口舌了。',
        'I speak with you today not as in ordinary conversation, so it behooves you to approach My words carefully, and even more so to reflect upon them more. You are even less willing to ponder more the disposition of God; seldom does anyone put effort into it. No wonder that you have no regard for God and no truth in your lives; how can such a group of people be deemed saints? Since you have so little understanding of this, I have no choice but to expend more breath.'
      ],
      [
        '神的性情对每一个人都是一个很抽象的问题，因为他的性情不同于一个人的性格。神自己有自己的所是所有，他的一切发表与流露代表他自己的实质。他的性情中所包括的有对人类的爱、对人类的抚慰，更有对人类望眼欲穿的了解。神的性情是最高权力的象征，任何一个人或任何一些人都不能也不可能搅扰他的工作与他的性情。',
        'The disposition of God is a subject that seems very abstract to everyone, for His disposition is unlike a human being’s personality. God Himself has His own possessions and being, and all that He expresses and is revealed from Him represents His own essence. His disposition encompasses His love for mankind, solace toward mankind, and even more, a thorough understanding of mankind. His disposition is the symbol of the highest power; no person or persons can disturb His work or His disposition.'
      ]
    ];
    const items = [
      { r: 0, za: '在这里', en: 'today', cmt: '原文【在这里】表空间（here），译文译作时间词 today（今天），请核查。', rep: '已核对，拟改为 here' },
      { r: 0, za: '慎重', en: 'carefully', cmt: '【慎重】侧重态度严肃郑重，carefully 偏小心仔细，请斟酌。' },
      { r: 0, za: '揣摩', en: 'reflect upon', cmt: '术语表【揣摩（神话）】用 ponder，reflect upon 对应【反省】。' },
      { r: 0, za: '多多斟酌一番', en: 'ponder more', cmt: '比较级 ponder more 预设原本已在揣摩，与后半句矛盾。' },
      { r: 0, za: '无真理', en: 'no truth', cmt: '【真理】须加定冠词 the truth，请看是否改为 are without the truth。' },
      { r: 0, za: '知之甚少', en: 'have so little understanding', cmt: '【知】译为 understanding 与本篇核心词【了解】撞词。' },
      { r: 1, za: '都是一个很抽象', en: 'seems very abstract', cmt: '原文为直接判断【是】，译文增译 seems 弱化为主观印象。' },
      { r: 1, za: '所是所有', en: 'possessions and being', cmt: '此处原文词序为【所是所有】，请确认是否调整为 being and possessions。' },
      { r: 1, za: '发表与流露', en: 'that He expresses and is revealed', cmt: '同一个 that 先作宾语后直接作被动主语，存在并列从句语法问题。' },
      { r: 1, za: '望眼欲穿', en: 'thorough', cmt: '【望眼欲穿】强调殷切期盼，thorough（透彻的）丢失了期盼之意。' }
    ];

    const buildCellXml = (rIdx, colIdx, text) => {
      let cuts = [0, text.length];
      const marks = [];
      items.forEach((it, idx) => {
        if (it.r !== rIdx) return;
        const sub = colIdx === 0 ? it.za : it.en;
        const pos = text.indexOf(sub);
        if (pos >= 0) {
          cuts.push(pos, pos + sub.length);
          marks.push({ a: pos, b: pos + sub.length, cid: idx + 1, isAnchor: colIdx === 0, hasRep: !!it.rep });
        }
      });
      cuts = Array.from(new Set(cuts)).sort((a, b) => a - b);
      let runsXml = '';
      for (let i = 0; i < cuts.length - 1; i++) {
        const a = cuts[i], b = cuts[i + 1];
        if (a >= b) continue;
        const seg = text.slice(a, b);
        const m = marks.find(x => x.a <= a && b <= x.b);
        if (m && m.isAnchor && m.a === a) {
          runsXml += `<w:commentRangeStart w:id="${m.cid}"/>` + (m.hasRep ? `<w:commentRangeStart w:id="101"/>` : '');
        }
        const hlXml = m ? `<w:rPr><w:highlight w:val="cyan"/></w:rPr>` : '';
        runsXml += `<w:r>${hlXml}<w:t xml:space="preserve">${seg}</w:t></w:r>`;
        if (m && m.isAnchor && m.b === b) {
          runsXml += `<w:commentRangeEnd w:id="${m.cid}"/><w:r><w:commentReference w:id="${m.cid}"/></w:r>`;
          if (m.hasRep) runsXml += `<w:commentRangeEnd w:id="101"/><w:r><w:commentReference w:id="101"/></w:r>`;
        }
      }
      return `<w:tc><w:p>${runsXml}</w:p></w:tc>`;
    };

    const trsXml = pairs.map((pr, ri) => `<w:tr>${buildCellXml(ri, 0, pr[0])}${buildCellXml(ri, 1, pr[1])}</w:tr>`).join('');
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${wNs}"><w:body><w:tbl>${trsXml}</w:tbl></w:body></w:document>`;
    const cmtsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${wNs}" xmlns:w14="${w14Ns}">` +
      items.map((it, idx) => `<w:comment w:id="${idx + 1}" w:author="审校" w:date="2026-11-18T01:00:00Z" w:initials="RV"><w:p w14:paraId="1000000${idx}"><w:r><w:t>${it.cmt}</w:t></w:r></w:p></w:comment>`).join('') +
      `<w:comment w:id="101" w:author="译者" w:date="2026-11-18T02:00:00Z" w:initials="TR"><w:p w14:paraId="20000001"><w:r><w:t>${items[0].rep}</w:t></w:r></w:p></w:comment></w:comments>`;
    const exXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="${w15Ns}">` +
      items.map((_, idx) => `<w15:commentEx w15:paraId="1000000${idx}" w15:done="0"/>`).join('') +
      `<w15:commentEx w15:paraId="20000001" w15:paraIdParent="10000000" w15:done="0"/></w15:commentsEx>`;
    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>`;
    const ctXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>`;

    const files = new Map([
      ['[Content_Types].xml', ctXml],
      ['word/document.xml', docXml],
      ['word/comments.xml', cmtsXml],
      ['word/commentsExtended.xml', exXml],
      ['word/_rels/document.xml.rels', relsXml]
    ]);
    const zipBytes = await window.ZipCodec.writeZip({ order: Array.from(files.keys()), files });
    await openDocxFile(new File([zipBytes], '示例双语批注稿.docx'));
  };

  window.exportCurrentDocx = exportCurrentDocx;
  window.undo = undo;
  window.redo = redo;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
