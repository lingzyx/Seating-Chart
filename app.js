/* global XLSX, ExcelJS */
(() => {
  'use strict';

  const state = {
    guests: [],
    seats: new Map(),
    rows: 13,
    cols: 8,
    verticalAisles: [],
    horizontalAisles: [],
  };

  const els = {
    fileInput: document.getElementById('fileInput'),
    exportBtn: document.getElementById('exportBtn'),
    clearGuestsBtn: document.getElementById('clearGuestsBtn'),
    clearSeatsBtn: document.getElementById('clearSeatsBtn'),
    applyLayoutBtn: document.getElementById('applyLayoutBtn'),
    rowInput: document.getElementById('rowInput'),
    colInput: document.getElementById('colInput'),
    verticalAislesInput: document.getElementById('verticalAislesInput'),
    horizontalAislesInput: document.getElementById('horizontalAislesInput'),
    guestSearch: document.getElementById('guestSearch'),
    guestList: document.getElementById('guestList'),
    guestDropZone: document.getElementById('guestDropZone'),
    unassignedCount: document.getElementById('unassignedCount'),
    totalCount: document.getElementById('totalCount'),
    seatGrid: document.getElementById('seatGrid'),
    guestDialog: document.getElementById('guestDialog'),
    guestForm: document.getElementById('guestForm'),
    editingGuestId: document.getElementById('editingGuestId'),
    editOrg: document.getElementById('editOrg'),
    editPerson: document.getElementById('editPerson'),
    editNote: document.getElementById('editNote'),
    cancelEditBtn: document.getElementById('cancelEditBtn'),
    closeDialogBtn: document.getElementById('closeDialogBtn'),
    toast: document.getElementById('toast'),
  };

  const ORG_ENDINGS = [
    '服務處', '辦公室', '事務所', '委員會', '管理處', '營業處', '代表會', '縣議會', '市議會',
    '縣政府', '市政府', '鄉公所', '鎮公所', '市公所', '區公所', '基金會', '協會',
    '大學', '高中', '國中', '國小', '醫院', '學校', '區處', '分處', '公司', '工會', '公會',
    '中心', '議會', '署', '局', '處', '所', '廠', '會'
  ];

  const TITLE_WORDS = [
    '縣長', '市長', '鄉長', '鎮長', '區長', '議員', '立委', '委員', '代表', '董事長', '總經理',
    '副總經理', '處長', '副處長', '廠長', '副廠長', '主任', '副主任', '課長', '組長', '股長',
    '秘書', '專員', '助理', '督導', '經理', '副理', '理事長', '總幹事', '會長', '校長', '里長'
  ];

  const COMMON_SURNAMES = new Set('王李張劉陳楊黃趙吳周徐孫馬朱胡郭何高林羅鄭梁謝宋唐許韓馮鄧曹彭曾蕭田董潘袁蔡蔣余于杜葉程魏蘇呂丁任沈姚盧姜崔鍾譚陸汪范金石廖賴侯邱方江白康游詹施洪簡藍顏莊詹溫傅呂柯盧阮魏歐陽上官司徒'.split(''));

  function seatKey(row, col) {
    return `${row}-${col}`;
  }

  function clampNumber(value, min, max, fallback) {
    const parsed = Number.parseInt(value, 10);
    if (Number.isNaN(parsed)) return fallback;
    return Math.min(max, Math.max(min, parsed));
  }

  function parseAisles(value, maxPosition) {
    if (!value.trim()) return [];
    return [...new Set(value
      .split(/[，,\s]+/)
      .map(item => Number.parseInt(item, 10))
      .filter(num => Number.isInteger(num) && num >= 1 && num < maxPosition))]
      .sort((a, b) => a - b);
  }

  function cleanText(value) {
    return String(value ?? '')
      .replace(/\r/g, '\n')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function normalizeHeader(header) {
    return cleanText(header).replace(/[\s_\-／/]/g, '').toLowerCase();
  }

  function detectHeaderIndexes(row) {
    const normalized = row.map(normalizeHeader);
    const findIndex = keywords => normalized.findIndex(cell => keywords.some(keyword => cell.includes(keyword)));
    const indexes = {
      org: findIndex(['機關單位', '機關', '單位', '公司', '部門']),
      name: findIndex(['姓名', '名字', '貴賓', '來賓', '賓客']),
      title: findIndex(['職稱', '稱謂', '職務']),
      note: findIndex(['備註', '說明', '備考', '註記']),
    };
    const hitCount = Object.values(indexes).filter(index => index >= 0).length;
    return hitCount >= 2 ? indexes : null;
  }

  function isLikelyNameAndTitle(text) {
    if (!text || text.length < 2) return false;
    const firstChar = text[0];
    const hasCommonSurname = COMMON_SURNAMES.has(firstChar);
    const earlyText = text.slice(0, 7);
    const hasTitle = TITLE_WORDS.some(title => earlyText.includes(title));
    return hasCommonSurname || hasTitle;
  }

  function detectOrgFromRaw(rawText) {
    const raw = cleanText(rawText).replace(/\s/g, '');
    if (!raw) return { org: '', personLine: '', note: '', raw: '' };

    const candidates = [];
    ORG_ENDINGS.forEach((ending, endingIndex) => {
      let start = raw.indexOf(ending);
      while (start >= 0) {
        const boundary = start + ending.length;
        const org = raw.slice(0, boundary);
        const rest = raw.slice(boundary);
        if (org.length >= 2 && rest.length >= 1) {
          let score = 0;
          score += Math.min(24, org.length);
          score += Math.max(0, 20 - endingIndex);
          if (ending.length >= 2) score += 18;
          if (isLikelyNameAndTitle(rest)) score += 40;
          if (rest.length >= 2 && rest.length <= 12) score += 12;
          if (org.length > 30) score -= 20;
          if (/^[長副助專]/.test(rest)) score -= 30;
          candidates.push({ org, personLine: rest, score, boundary });
        }
        start = raw.indexOf(ending, start + 1);
      }
    });

    if (!candidates.length) {
      return { org: '', personLine: raw, note: '', raw };
    }

    candidates.sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      return a.boundary - b.boundary;
    });

    const best = candidates[0];
    return { org: best.org, personLine: best.personLine, note: '', raw };
  }

  function composePersonLine(name, title, note) {
    const cleanName = cleanText(name).replace(/\s/g, '');
    const cleanTitle = cleanText(title).replace(/\s/g, '');
    const cleanNote = cleanText(note);

    let line = cleanName;
    if (cleanTitle) {
      if (!cleanName) {
        line = cleanTitle;
      } else if (cleanName.includes(cleanTitle)) {
        line = cleanName;
      } else if (/^[\u4e00-\u9fff]{2,}$/.test(cleanName)) {
        line = `${cleanName[0]}${cleanTitle}${cleanName.slice(1)}`;
      } else {
        line = `${cleanTitle}${cleanName}`;
      }
    }

    return cleanNote ? `${line}（${cleanNote}）` : line;
  }

  function makeGuest(data, index) {
    return {
      id: `guest-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 8)}`,
      org: data.org || '',
      personLine: data.personLine || '',
      note: data.note || '',
      raw: data.raw || '',
    };
  }

  function rowsToGuests(rows) {
    const filteredRows = rows
      .map(row => row.map(cleanText))
      .filter(row => row.some(Boolean));

    if (!filteredRows.length) return [];

    const headerIndexes = detectHeaderIndexes(filteredRows[0]);
    const dataRows = headerIndexes ? filteredRows.slice(1) : filteredRows;

    return dataRows.map((row, index) => {
      if (headerIndexes) {
        const org = headerIndexes.org >= 0 ? row[headerIndexes.org] : '';
        const name = headerIndexes.name >= 0 ? row[headerIndexes.name] : '';
        const title = headerIndexes.title >= 0 ? row[headerIndexes.title] : '';
        const note = headerIndexes.note >= 0 ? row[headerIndexes.note] : '';
        const hasStructuredData = org || name || title || note;
        if (hasStructuredData) {
          const personLine = composePersonLine(name, title, note);
          const raw = [org, name, title, note].filter(Boolean).join(' ');
          return makeGuest({ org, personLine, note, raw }, index);
        }
      }

      const compact = row.filter(Boolean).join('');
      return makeGuest(detectOrgFromRaw(compact), index);
    }).filter(guest => guest.org || guest.personLine || guest.raw);
  }

  async function handleFileUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) throw new Error('找不到工作表');

      const worksheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
      const guests = rowsToGuests(rows);

      state.guests = guests;
      state.seats.clear();
      renderAll();
      showToast(`已讀取 ${guests.length} 位賓客`);
    } catch (error) {
      console.error(error);
      showToast('讀取 Excel 失敗，請確認檔案格式');
    } finally {
      event.target.value = '';
    }
  }

  function guestById(id) {
    return state.guests.find(guest => guest.id === id);
  }

  function assignedGuestIds() {
    return new Set(state.seats.values());
  }

  function findSeatByGuestId(guestId) {
    for (const [key, id] of state.seats.entries()) {
      if (id === guestId) return key;
    }
    return null;
  }

  function getUnassignedGuests() {
    const assigned = assignedGuestIds();
    return state.guests.filter(guest => !assigned.has(guest.id));
  }

  function renderGuestList() {
    const query = cleanText(els.guestSearch.value).toLowerCase();
    const unassigned = getUnassignedGuests();
    const visibleGuests = unassigned.filter(guest => {
      const haystack = `${guest.org} ${guest.personLine} ${guest.note} ${guest.raw}`.toLowerCase();
      return haystack.includes(query);
    });

    els.unassignedCount.textContent = String(unassigned.length);
    els.totalCount.textContent = String(state.guests.length);
    els.guestList.innerHTML = '';
    els.guestList.classList.toggle('empty', visibleGuests.length === 0);

    if (!state.guests.length) {
      els.guestList.innerHTML = `
        <div class="empty-state">
          <div class="empty-icon">📄</div>
          <p>請先上傳 Excel 名單</p>
          <small>支援有欄位的名單，或只有一欄完整文字的名單。</small>
        </div>
      `;
      return;
    }

    if (!visibleGuests.length) {
      els.guestList.innerHTML = `
        <div class="empty-state">
          <div class="empty-icon">✅</div>
          <p>${unassigned.length ? '找不到符合搜尋的賓客' : '所有賓客都已安排'}</p>
          <small>${unassigned.length ? '請換一個關鍵字搜尋。' : '可從座位點選清除，讓賓客回到名單。'}</small>
        </div>
      `;
      return;
    }

    const fragment = document.createDocumentFragment();
    visibleGuests.forEach(guest => fragment.appendChild(createGuestCard(guest, 'list')));
    els.guestList.appendChild(fragment);
  }

  function createGuestCard(guest, source, fromKey = '') {
    const card = document.createElement('article');
    card.className = 'guest-card';
    card.draggable = true;
    card.dataset.guestId = guest.id;
    card.innerHTML = `
      <div class="guest-org">${escapeHtml(guest.org || '未辨識機關單位')}</div>
      <div class="guest-person">${escapeHtml(guest.personLine || guest.raw || '未命名賓客')}</div>
      <div class="card-actions">
        <button class="tiny-button" type="button" data-action="edit">編輯</button>
      </div>
      ${guest.note ? `<div class="guest-note">${escapeHtml(guest.note)}</div>` : ''}
    `;

    card.addEventListener('dragstart', event => {
      card.classList.add('dragging');
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', JSON.stringify({ source, guestId: guest.id, fromKey }));
    });
    card.addEventListener('dragend', () => card.classList.remove('dragging'));
    card.querySelector('[data-action="edit"]').addEventListener('click', event => {
      event.stopPropagation();
      openGuestDialog(guest.id);
    });
    return card;
  }

  function makeGridTemplateColumns() {
    const columns = [];
    for (let col = 1; col <= state.cols; col += 1) {
      columns.push('112px');
      if (state.verticalAisles.includes(col)) columns.push('28px');
    }
    return columns.join(' ');
  }

  function visualColumnCount() {
    return state.cols + state.verticalAisles.length;
  }

  function renderSeatGrid() {
    els.seatGrid.innerHTML = '';
    els.seatGrid.style.gridTemplateColumns = makeGridTemplateColumns();

    const fragment = document.createDocumentFragment();
    for (let row = 1; row <= state.rows; row += 1) {
      for (let col = 1; col <= state.cols; col += 1) {
        fragment.appendChild(createSeatCell(row, col));
        if (state.verticalAisles.includes(col)) {
          const aisle = document.createElement('div');
          aisle.className = 'aisle-cell vertical';
          aisle.textContent = '走道';
          fragment.appendChild(aisle);
        }
      }
      if (state.horizontalAisles.includes(row)) {
        const aisle = document.createElement('div');
        aisle.className = 'aisle-cell horizontal';
        aisle.textContent = '走道';
        aisle.style.gridColumn = `1 / span ${visualColumnCount()}`;
        fragment.appendChild(aisle);
      }
    }
    els.seatGrid.appendChild(fragment);
  }

  function createSeatCell(row, col) {
    const key = seatKey(row, col);
    const guestId = state.seats.get(key);
    const guest = guestId ? guestById(guestId) : null;
    const cell = document.createElement('div');
    cell.className = `seat-cell${guest ? ' assigned' : ''}`;
    cell.dataset.key = key;
    cell.innerHTML = `
      <div class="seat-number">${row}-${col}</div>
      <div class="seat-content">
        ${guest ? `<div class="seat-org">${escapeHtml(guest.org || '未辨識機關單位')}</div><div class="seat-person">${escapeHtml(guest.personLine || guest.raw || '')}</div>` : '<span class="seat-empty">拖拉至此</span>'}
      </div>
      ${guest ? `<div class="seat-actions"><button class="seat-action-button" type="button" data-action="edit" title="編輯">✎</button><button class="seat-action-button" type="button" data-action="clear" title="清除座位">×</button></div>` : ''}
    `;

    cell.addEventListener('dragover', event => {
      event.preventDefault();
      cell.classList.add('drag-over');
    });
    cell.addEventListener('dragleave', () => cell.classList.remove('drag-over'));
    cell.addEventListener('drop', event => {
      event.preventDefault();
      cell.classList.remove('drag-over');
      handleSeatDrop(key, event.dataTransfer.getData('text/plain'));
    });

    if (guest) {
      cell.draggable = true;
      cell.addEventListener('dragstart', event => {
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('text/plain', JSON.stringify({ source: 'seat', guestId: guest.id, fromKey: key }));
      });
      cell.querySelector('[data-action="clear"]').addEventListener('click', event => {
        event.stopPropagation();
        state.seats.delete(key);
        renderAll();
        showToast('已清除座位，賓客回到未安排名單');
      });
      cell.querySelector('[data-action="edit"]').addEventListener('click', event => {
        event.stopPropagation();
        openGuestDialog(guest.id);
      });
    }

    return cell;
  }

  function handleSeatDrop(targetKey, payloadText) {
    const payload = parseDragPayload(payloadText);
    if (!payload) return;

    const { guestId, source, fromKey } = payload;
    const targetGuestId = state.seats.get(targetKey);
    const currentSeatKey = findSeatByGuestId(guestId);

    if (currentSeatKey === targetKey) return;

    if (source === 'seat' && fromKey) {
      state.seats.delete(fromKey);
      if (targetGuestId) state.seats.set(fromKey, targetGuestId);
      state.seats.set(targetKey, guestId);
    } else {
      if (currentSeatKey) state.seats.delete(currentSeatKey);
      state.seats.set(targetKey, guestId);
    }

    renderAll();
  }

  function handleDropToGuestList(payloadText) {
    const payload = parseDragPayload(payloadText);
    if (!payload) return;
    const currentSeatKey = findSeatByGuestId(payload.guestId);
    if (currentSeatKey) {
      state.seats.delete(currentSeatKey);
      renderAll();
      showToast('已移回未安排名單');
    }
  }

  function parseDragPayload(text) {
    try {
      const payload = JSON.parse(text);
      if (!payload || !payload.guestId || !guestById(payload.guestId)) return null;
      return payload;
    } catch {
      return null;
    }
  }

  function applyLayout() {
    const rows = clampNumber(els.rowInput.value, 1, 80, state.rows);
    const cols = clampNumber(els.colInput.value, 1, 80, state.cols);
    const verticalAisles = parseAisles(els.verticalAislesInput.value, cols);
    const horizontalAisles = parseAisles(els.horizontalAislesInput.value, rows);

    state.rows = rows;
    state.cols = cols;
    state.verticalAisles = verticalAisles;
    state.horizontalAisles = horizontalAisles;

    for (const key of [...state.seats.keys()]) {
      const [row, col] = key.split('-').map(Number);
      if (row > rows || col > cols) state.seats.delete(key);
    }

    els.rowInput.value = String(rows);
    els.colInput.value = String(cols);
    els.verticalAislesInput.value = verticalAisles.join(', ');
    els.horizontalAislesInput.value = horizontalAisles.join(', ');

    renderAll();
    showToast('已套用座位配置');
  }

  function clearSeats() {
    if (!state.seats.size) {
      showToast('目前沒有已安排座位');
      return;
    }
    if (!window.confirm('確定要清空所有座位安排嗎？賓客名單會保留。')) return;
    state.seats.clear();
    renderAll();
    showToast('已清空所有座位');
  }

  function clearGuests() {
    if (!state.guests.length) {
      showToast('目前沒有名單可清空');
      return;
    }
    if (!window.confirm('確定要清空賓客名單與所有座位安排嗎？')) return;
    state.guests = [];
    state.seats.clear();
    renderAll();
    showToast('已清空名單與座位');
  }

  function openGuestDialog(guestId) {
    const guest = guestById(guestId);
    if (!guest) return;
    els.editingGuestId.value = guest.id;
    els.editOrg.value = guest.org;
    els.editPerson.value = guest.personLine;
    els.editNote.value = guest.note;
    els.guestDialog.showModal();
    els.editOrg.focus();
  }

  function closeGuestDialog() {
    els.guestDialog.close();
  }

  function saveGuestEdit(event) {
    event.preventDefault();
    const guest = guestById(els.editingGuestId.value);
    if (!guest) return;

    guest.org = cleanText(els.editOrg.value);
    guest.personLine = cleanText(els.editPerson.value);
    guest.note = cleanText(els.editNote.value);
    guest.raw = [guest.org, guest.personLine, guest.note].filter(Boolean).join(' ');

    closeGuestDialog();
    renderAll();
    showToast('已更新賓客資料');
  }

  function renderAll() {
    renderGuestList();
    renderSeatGrid();
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function showToast(message) {
    els.toast.textContent = message;
    els.toast.classList.add('show');
    window.clearTimeout(showToast.timer);
    showToast.timer = window.setTimeout(() => els.toast.classList.remove('show'), 2400);
  }

  function excelCellAddress(row, col) {
    return { row, col };
  }

  async function exportExcel() {
    if (!state.rows || !state.cols) {
      showToast('請先設定座位配置');
      return;
    }

    if (typeof ExcelJS === 'undefined') {
      showToast('Excel 匯出套件尚未載入，請檢查網路後重試');
      return;
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = '禮堂座位安排工具';
    workbook.created = new Date();
    const worksheet = workbook.addWorksheet('座位表', {
      properties: { defaultRowHeight: 54 },
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    const visualCols = buildVisualColumns();
    visualCols.forEach((item, index) => {
      const column = worksheet.getColumn(index + 1);
      column.width = item.type === 'aisle' ? 4 : 18;
    });

    let excelRow = 1;
    for (let row = 1; row <= state.rows; row += 1) {
      let excelCol = 1;
      for (let col = 1; col <= state.cols; col += 1) {
        writeSeatCell(worksheet, excelRow, excelCol, row, col);
        excelCol += 1;
        if (state.verticalAisles.includes(col)) {
          worksheet.getCell(excelRow, excelCol).value = '';
          excelCol += 1;
        }
      }
      worksheet.getRow(excelRow).height = 54;
      excelRow += 1;

      if (state.horizontalAisles.includes(row)) {
        worksheet.getRow(excelRow).height = 18;
        excelRow += 1;
      }
    }

    worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 0 }];

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    downloadBlob(blob, `禮堂座位表_${formatDate(new Date())}.xlsx`);
    showToast('已產出 Excel 座位表');
  }

  function buildVisualColumns() {
    const columns = [];
    for (let col = 1; col <= state.cols; col += 1) {
      columns.push({ type: 'seat', col });
      if (state.verticalAisles.includes(col)) columns.push({ type: 'aisle' });
    }
    return columns;
  }

  function writeSeatCell(worksheet, excelRow, excelCol, seatRow, seatCol) {
    const cell = worksheet.getCell(excelRow, excelCol);
    const key = seatKey(seatRow, seatCol);
    const guest = guestById(state.seats.get(key));

    if (guest) {
      const org = guest.org || '';
      const person = guest.personLine || guest.raw || '';
      if (org && person) {
        cell.value = {
          richText: [
            { text: org, font: { name: 'Microsoft JhengHei', bold: true, size: 11 } },
            { text: `\n${person}`, font: { name: 'Microsoft JhengHei', bold: false, size: 11 } },
          ]
        };
      } else {
        cell.value = org || person;
        cell.font = { name: 'Microsoft JhengHei', size: 11, bold: Boolean(org && !person) };
      }
    } else {
      cell.value = '';
    }

    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF9FB4AD' } },
      left: { style: 'thin', color: { argb: 'FF9FB4AD' } },
      bottom: { style: 'thin', color: { argb: 'FF9FB4AD' } },
      right: { style: 'thin', color: { argb: 'FF9FB4AD' } },
    };
    // 不設定 fill，讓 Excel 座位儲存格維持無底色。
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function formatDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    const hh = String(date.getHours()).padStart(2, '0');
    const mm = String(date.getMinutes()).padStart(2, '0');
    return `${y}${m}${d}_${hh}${mm}`;
  }

  function setupDropZone() {
    els.guestDropZone.addEventListener('dragover', event => {
      event.preventDefault();
      els.guestDropZone.classList.add('drag-over');
    });
    els.guestDropZone.addEventListener('dragleave', () => els.guestDropZone.classList.remove('drag-over'));
    els.guestDropZone.addEventListener('drop', event => {
      event.preventDefault();
      els.guestDropZone.classList.remove('drag-over');
      handleDropToGuestList(event.dataTransfer.getData('text/plain'));
    });
  }

  function bindEvents() {
    els.fileInput.addEventListener('change', handleFileUpload);
    els.exportBtn.addEventListener('click', exportExcel);
    els.applyLayoutBtn.addEventListener('click', applyLayout);
    els.clearSeatsBtn.addEventListener('click', clearSeats);
    els.clearGuestsBtn.addEventListener('click', clearGuests);
    els.guestSearch.addEventListener('input', renderGuestList);
    els.guestForm.addEventListener('submit', saveGuestEdit);
    els.cancelEditBtn.addEventListener('click', closeGuestDialog);
    els.closeDialogBtn.addEventListener('click', closeGuestDialog);
    setupDropZone();
  }

  bindEvents();
  renderAll();
})();
