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
    workspace: document.getElementById('workspace'),
    toggleGuestPanelBtn: document.getElementById('toggleGuestPanelBtn'),
    expandGuestPanelBtn: document.getElementById('expandGuestPanelBtn'),
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

  const SPECIAL_ORG_LABELS = [
    // 這類不是正式機關結尾詞，但在座位表上常作為分類名稱，匯出時也要放第一行粗體。
    '處長親友', '退休長官'
  ];

  const ORG_ENDINGS = [
    // 較長、較具體的組織結尾要放前面，避免被「公司、廠、會」等短詞過早切斷。
    '廠商聯合總會', '工業區廠商聯合總會', '聯合總會', '勞工董事會', '董事會',
    '工程服務社', '服務社', '辦事處', '服務處', '辦公室', '事務所', '委員會', '管理處', '營業處',
    '第十分會', '第九分會', '第八分會', '第七分會', '第六分會', '第五分會', '第四分會', '第三分會', '第二分會', '第一分會', '分會',
    '代表會', '縣議會', '市議會', '縣政府', '市政府', '鄉公所', '鎮公所', '市公所',
    '區公所', '基金會', '協會', '總會', '大學', '高中', '國中', '國小', '醫院', '學校',
    '區處', '分處', '公司', '工會', '公會', '中心', '議會', '署', '局', '處', '所', '廠', '會'
  ];

  const TITLE_WORDS = [
    '副理事長', '常務理事', '常務監事', '主任委員', '董事長', '副董事長', '總經理', '副總經理',
    '縣長', '市長', '鄉長', '鎮長', '區長', '議員', '立委', '委員', '代表',
    '處長', '副處長', '廠長', '副廠長', '主任', '副主任', '課長', '組長', '股長',
    '秘書', '專員', '助理', '督導', '經理', '副理', '主委', '理事長', '總幹事', '會長', '校長', '里長'
  ];

  const COMMON_SURNAMES = new Set('王李張劉陳楊黃趙吳周徐孫馬朱胡郭何高林羅鄭梁謝宋唐許韓馮鄧曹彭曾蕭田董潘袁蔡蔣余于杜葉程魏蘇呂丁任沈姚盧姜崔鍾譚陸汪范金石廖賴侯邱方江白康游詹施洪簡藍顏莊詹溫傅呂柯盧阮魏歐陽上官司徒'.split(''));


  const ORG_COLOR_PALETTE = [
    { bg: '#f0f8f5', border: '#b9dfd2', accent: '#6fb7a3' },
    { bg: '#f5f8ee', border: '#d2e1ad', accent: '#9fbd67' },
    { bg: '#fff7ed', border: '#efd4aa', accent: '#d7a45f' },
    { bg: '#f1f7fb', border: '#b9d7e7', accent: '#6aa7c7' },
    { bg: '#f8f3fb', border: '#d9c4e8', accent: '#ad81c6' },
    { bg: '#fff3f5', border: '#edc2cb', accent: '#d88698' },
    { bg: '#f2f7ff', border: '#c4d7f1', accent: '#7fa8dc' },
    { bg: '#f7f5ef', border: '#ded3b7', accent: '#b9a16a' },
    { bg: '#eef8f8', border: '#acdcdc', accent: '#63b4b4' },
    { bg: '#f9f2ef', border: '#e2c7b9', accent: '#c48b73' },
    { bg: '#f3f7f0', border: '#c6dcbd', accent: '#83b071' },
    { bg: '#f4f4fb', border: '#c9c9ec', accent: '#8f8bd1' },
  ];

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


  function normalizeOrgForColor(org) {
    return cleanText(org).replace(/[\s　]/g, '').replace(/[()]/g, match => (match === '(' ? '（' : '）'));
  }

  function hashString(value) {
    let hash = 0;
    for (let index = 0; index < value.length; index += 1) {
      hash = ((hash << 5) - hash) + value.charCodeAt(index);
      hash |= 0;
    }
    return Math.abs(hash);
  }

  function orgColorStyle(org) {
    const normalizedOrg = normalizeOrgForColor(org);
    if (!normalizedOrg) return null;
    return ORG_COLOR_PALETTE[hashString(normalizedOrg) % ORG_COLOR_PALETTE.length];
  }

  function applyOrgColorStyle(element, org) {
    const color = orgColorStyle(org);
    if (!color) return;
    element.style.setProperty('--guest-group-bg', color.bg);
    element.style.setProperty('--guest-group-border', color.border);
    element.style.setProperty('--guest-group-accent', color.accent);
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
    const earlyText = text.slice(0, 10);
    const hasTitle = TITLE_WORDS.some(title => earlyText.includes(title));
    return hasCommonSurname || hasTitle;
  }

  function startsWithTitle(text) {
    return TITLE_WORDS.some(title => text.startsWith(title));
  }

  function looksLikeOrgContinuation(text) {
    if (!text) return false;
    return /^[商業工產區聯合總董監事會務服處局所中心公協勞民分第0-9０-９一二三四五六七八九十]/.test(text);
  }

  function includeTrailingParentheses(raw, boundary) {
    let nextBoundary = boundary;
    let hasParentheses = false;

    while (nextBoundary < raw.length) {
      const open = raw[nextBoundary];
      const close = open === '（' ? '）' : open === '(' ? ')' : '';
      if (!close) break;

      const closeIndex = raw.indexOf(close, nextBoundary + 1);
      if (closeIndex < 0) break;

      nextBoundary = closeIndex + 1;
      hasParentheses = true;
    }

    return {
      boundary: nextBoundary,
      hasParentheses,
    };
  }

  function detectOrgFromRaw(rawText) {
    const raw = cleanText(rawText).replace(/\s/g, '');
    if (!raw) return { org: '', personLine: '', note: '', raw: '' };

    const specialLabel = SPECIAL_ORG_LABELS.find(label => raw.startsWith(label));
    if (specialLabel) {
      return { org: specialLabel, personLine: raw.slice(specialLabel.length), note: '', raw };
    }

    const candidates = [];
    ORG_ENDINGS.forEach((ending, endingIndex) => {
      let start = raw.indexOf(ending);
      while (start >= 0) {
        const initialBoundary = start + ending.length;
        const extended = includeTrailingParentheses(raw, initialBoundary);
        const boundary = extended.boundary;
        const org = raw.slice(0, boundary);
        const rest = raw.slice(boundary);
        if (org.length >= 2 && (rest.length >= 1 || extended.hasParentheses)) {
          let score = 0;
          score += Math.min(40, org.length * 2);
          score += Math.max(0, 28 - endingIndex);
          score += ending.length * 12;
          if (extended.hasParentheses) score += 45;
          if (startsWithTitle(rest)) score += 55;
          else if (isLikelyNameAndTitle(rest)) score += 40;
          else if (!rest) score += 18;
          else score -= 35;
          if (looksLikeOrgContinuation(rest)) score -= 55;
          if (rest.length >= 2 && rest.length <= 12) score += 12;
          if (org.length > 30) score -= 20;
          candidates.push({ org, personLine: rest, score, boundary, endingLength: ending.length });
        }
        start = raw.indexOf(ending, start + 1);
      }
    });

    if (!candidates.length) {
      return { org: '', personLine: raw, note: '', raw };
    }

    candidates.sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      if (b.endingLength !== a.endingLength) return b.endingLength - a.endingLength;
      return b.boundary - a.boundary;
    });

    const best = candidates[0];
    return { org: best.org, personLine: best.personLine, note: '', raw };
  }

  function refineStructuredOrg(org, personLine) {
    const cleanOrg = cleanText(org).replace(/\s/g, '');
    const cleanPersonLine = cleanText(personLine).replace(/\s/g, '');
    if (!cleanOrg || !cleanPersonLine) return { org: cleanOrg, personLine: cleanPersonLine };

    const detected = detectOrgFromRaw(`${cleanOrg}${cleanPersonLine}`);
    if (detected.org && detected.org.startsWith(cleanOrg) && detected.org.length > cleanOrg.length && detected.personLine) {
      return { org: detected.org, personLine: detected.personLine };
    }

    return { org: cleanOrg, personLine: cleanPersonLine };
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
          const refined = refineStructuredOrg(org, personLine);
          const raw = [org, name, title, note].filter(Boolean).join(' ');
          return makeGuest({ org: refined.org, personLine: refined.personLine, note, raw }, index);
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
    card.className = `guest-card${guest.org ? ' has-org-color' : ''}`;
    applyOrgColorStyle(card, guest.org);
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
      columns.push('96px');
      if (state.verticalAisles.includes(col)) columns.push('22px');
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

    createSeatWorksheet(workbook, '台下視角', false);
    createSeatWorksheet(workbook, '台上視角', true);

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    downloadBlob(blob, `禮堂座位表_雙視角_${formatDate(new Date())}.xlsx`);
    showToast('已產出含台下與台上視角的 Excel 座位表');
  }

  function createSeatWorksheet(workbook, sheetName, rotateForStageView) {
    const worksheet = workbook.addWorksheet(sheetName, {
      properties: { defaultRowHeight: 54 },
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    const visualRows = buildVisualSeatLayout(rotateForStageView);
    const visualColCount = visualRows[0]?.length || 0;

    for (let colIndex = 0; colIndex < visualColCount; colIndex += 1) {
      const hasSeat = visualRows.some(row => row[colIndex]?.type === 'seat');
      worksheet.getColumn(colIndex + 1).width = hasSeat ? 18 : 4;
    }

    visualRows.forEach((layoutRow, rowIndex) => {
      const excelRow = rowIndex + 1;
      const rowHasSeat = layoutRow.some(item => item.type === 'seat');
      worksheet.getRow(excelRow).height = rowHasSeat ? 54 : 18;

      layoutRow.forEach((item, colIndex) => {
        const excelCol = colIndex + 1;
        if (item.type === 'seat') {
          writeSeatCell(worksheet, excelRow, excelCol, item.row, item.col);
        } else {
          worksheet.getCell(excelRow, excelCol).value = '';
        }
      });
    });

    worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 0 }];
    return worksheet;
  }

  function buildVisualSeatLayout(rotateForStageView = false) {
    const rows = [];
    const visualColCount = state.cols + state.verticalAisles.length;

    for (let row = 1; row <= state.rows; row += 1) {
      const layoutRow = [];
      for (let col = 1; col <= state.cols; col += 1) {
        layoutRow.push({ type: 'seat', row, col });
        if (state.verticalAisles.includes(col)) layoutRow.push({ type: 'aisle' });
      }
      rows.push(layoutRow);

      if (state.horizontalAisles.includes(row)) {
        rows.push(Array.from({ length: visualColCount }, () => ({ type: 'aisle' })));
      }
    }

    if (!rotateForStageView) return rows;
    return rows
      .slice()
      .reverse()
      .map(layoutRow => layoutRow.slice().reverse());
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

  function setGuestPanelCollapsed(collapsed) {
    if (!els.workspace) return;
    els.workspace.classList.toggle('guest-collapsed', collapsed);
    if (els.toggleGuestPanelBtn) {
      els.toggleGuestPanelBtn.textContent = collapsed ? '展開' : '縮小';
      els.toggleGuestPanelBtn.setAttribute('aria-expanded', String(!collapsed));
    }
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
    els.toggleGuestPanelBtn?.addEventListener('click', () => setGuestPanelCollapsed(true));
    els.expandGuestPanelBtn?.addEventListener('click', () => setGuestPanelCollapsed(false));
    els.guestSearch.addEventListener('input', renderGuestList);
    els.guestForm.addEventListener('submit', saveGuestEdit);
    els.cancelEditBtn.addEventListener('click', closeGuestDialog);
    els.closeDialogBtn.addEventListener('click', closeGuestDialog);
    setupDropZone();
  }

  bindEvents();
  renderAll();
})();
