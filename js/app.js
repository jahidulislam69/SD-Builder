// UI Controller, Full LocalStorage Persistence, and Focused Workspace
(function (root) {
  const {
    APP_CONSTANTS,
    SESSION_PLACEHOLDERS,
    resolveWeekDates,
    calculateDayAllocation,
    generateReportRows,
    exportReportToExcel,
    parseAuditWorkbook,
    syncFromGoogleSheet,
    DEFAULT_SHEET_URL
  } = root.SD;

  const STORAGE_KEY = 'sd_report_state_v1';

  function getDefaultSunday() {
    const d = new Date();
    const diff = d.getDate() - d.getDay();
    const sunday = new Date(d.getFullYear(), d.getMonth(), diff, 12, 0, 0);
    const pad = (n) => String(n).padStart(2, '0');
    return `${sunday.getFullYear()}-${pad(sunday.getMonth() + 1)}-${pad(sunday.getDate())}`;
  }

  function createDefaultDays(sundayStr) {
    const dates = resolveWeekDates(sundayStr);
    return dates.map((dateMeta, idx) => ({
      index: idx,
      type: 'NORMAL',
      voiceCount: idx === 0 ? 50 : 40,
      nonVoiceCount: idx === 0 ? 0 : 10,
      sessions: [],
      dateMeta
    }));
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && parsed.startDate && Array.isArray(parsed.days) && parsed.days.length === 5) {
          // Rehydrate Date objects in dateMeta from startDate
          const dates = resolveWeekDates(parsed.startDate);
          parsed.days.forEach((day, idx) => {
            day.dateMeta = dates[idx];
            day.sessions = Array.isArray(day.sessions) ? day.sessions : [];
          });
          return {
            profile: {
              name: parsed.profile?.name ?? 'Jahidul Islam Shuvo',
              qcBy: parsed.profile?.qcBy ?? APP_CONSTANTS.QC_BY_DEFAULT,
              unit: APP_CONSTANTS.UNIT
            },
            startDate: parsed.startDate,
            activeDayIndex: Math.min(4, Math.max(0, parsed.activeDayIndex || 0)),
            previewOpen: Boolean(parsed.previewOpen),
            days: parsed.days
          };
        }
      }
    } catch (e) {
      // Fallback to fresh state if storage is corrupted
    }

    const startDate = getDefaultSunday();
    return {
      profile: {
        name: 'Jahidul Islam Shuvo',
        qcBy: APP_CONSTANTS.QC_BY_DEFAULT,
        unit: APP_CONSTANTS.UNIT
      },
      startDate,
      activeDayIndex: 0,
      previewOpen: false,
      days: createDefaultDays(startDate)
    };
  }

  const state = loadState();

  function saveState() {
    try {
      // Strip Date objects before serializing; dateMeta is rehydrated from startDate
      const serializable = {
        profile: state.profile,
        startDate: state.startDate,
        activeDayIndex: state.activeDayIndex,
        previewOpen: state.previewOpen,
        days: state.days.map((d) => ({
          index: d.index,
          type: d.type,
          voiceCount: d.voiceCount,
          nonVoiceCount: d.nonVoiceCount,
          sessions: d.sessions
        }))
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(serializable));
      flashSaveIndicator();
    } catch (e) {
      // Ignore storage quota errors
    }
  }

  let saveTimer = null;
  function flashSaveIndicator() {
    const el = document.getElementById('saveIndicator');
    if (!el) return;
    el.textContent = '● Saving...';
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      el.textContent = '● Saved locally';
    }, 350);
  }

  // DOM References
  const elProcessorName = document.getElementById('processorName');
  const elQcBy = document.getElementById('qcBy');
  const elStartDate = document.getElementById('startDate');
  const elWeekRangeBadge = document.getElementById('weekRangeBadge');
  const elBtnPrevWeek = document.getElementById('btnPrevWeek');
  const elBtnNextWeek = document.getElementById('btnNextWeek');
  const elBtnResetWeek = document.getElementById('btnResetWeek');
  const elBtnImportAudit = document.getElementById('btnImportAudit');
  const elAuditFileInput = document.getElementById('auditFileInput');
  const elBtnSyncSheet = document.getElementById('btnSyncSheet');
  const elSyncFeedback = document.getElementById('syncFeedback');
  const elSyncFeedbackText = document.getElementById('syncFeedbackText');
  const elBtnDismissFeedback = document.getElementById('btnDismissFeedback');
  const elNameSuggestions = document.getElementById('processorNameSuggestions');
  const elBtnDemoShuvo = document.getElementById('btnDemoShuvo');
  const elBtnDemoRony = document.getElementById('btnDemoRony');
  const elBtnExport = document.getElementById('btnExport');
  const elDayStrip = document.getElementById('dayStrip');
  const elWorkspace = document.getElementById('activeDayWorkspace');
  const elPreviewToggleBar = document.getElementById('previewToggleBar');
  const elPreviewBody = document.getElementById('previewBody');
  const elPreviewChevron = document.getElementById('previewChevron');
  const elRowCountBadge = document.getElementById('rowCountBadge');
  const elPreviewTableHead = document.getElementById('previewTableHead');
  const elPreviewTableBody = document.getElementById('previewTableBody');

  function formatDateShort(dateObj) {
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = dateObj.toLocaleString('en-US', { month: 'short' });
    const year = dateObj.getFullYear();
    return `${day} ${month} ${year}`;
  }

  function formatDateCompact(dateObj) {
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = dateObj.toLocaleString('en-US', { month: 'short' });
    return `${day} ${month}`;
  }

  function renderWeekRange() {
    const first = state.days[0].dateMeta.completionDate;
    const last = state.days[4].dateMeta.completionDate;
    elWeekRangeBadge.textContent = `(${formatDateCompact(first)} – ${formatDateShort(last)})`;
  }

  function renderDayStrip() {
    elDayStrip.innerHTML = state.days.map((day, idx) => {
      const alloc = calculateDayAllocation(day);
      const isActive = idx === state.activeDayIndex;
      const isLeave = day.type === 'SICK_LEAVE' || day.type === 'CASUAL_LEAVE';
      const isHalf = day.type === 'HALF_DAY_LEAVE';

      let pillClass = 'pill-normal';
      let pillText = `${alloc.totalDayTime}m`;
      if (isLeave) {
        pillClass = 'pill-leave';
        pillText = day.type === 'SICK_LEAVE' ? 'Sick Leave' : 'Casual Leave';
      } else if (!alloc.isBalanced) {
        pillClass = 'pill-warn';
        pillText = `${alloc.totalDayTime}m`;
      } else if (isHalf) {
        pillClass = 'pill-half';
        pillText = 'Half Day';
      }

      const summaryText = isLeave
        ? '0 audits'
        : `${day.voiceCount}V / ${day.nonVoiceCount}NV${day.sessions.length ? ` · ${day.sessions.length} mtg` : ''}`;

      return `
        <button type="button" class="day-tab ${isActive ? 'active' : ''}" onclick="window.selectDay(${idx})">
          <div class="day-tab-top">
            <span class="day-tab-name">${day.dateMeta.dayName.slice(0, 3)}</span>
            <span class="day-tab-date">${formatDateCompact(day.dateMeta.completionDate)}</span>
          </div>
          <div class="day-tab-stats">
            <span style="color: var(--text-secondary);">${summaryText}</span>
            <span class="pill ${pillClass}">${pillText}</span>
          </div>
        </button>
      `;
    }).join('');
  }

  function renderActiveWorkspace() {
    const idx = state.activeDayIndex;
    const day = state.days[idx];
    const alloc = calculateDayAllocation(day);
    const meta = day.dateMeta;
    const isLeave = day.type === 'SICK_LEAVE' || day.type === 'CASUAL_LEAVE';
    const isHalfDay = day.type === 'HALF_DAY_LEAVE';

    const modeButtons = [
      { key: 'NORMAL', label: 'Normal Day' },
      { key: 'HALF_DAY_LEAVE', label: 'Half Day Leave' },
      { key: 'SICK_LEAVE', label: 'Sick Leave' },
      { key: 'CASUAL_LEAVE', label: 'Casual Leave' }
    ].map((m) => `
      <button type="button" class="seg-btn ${day.type === m.key ? 'active' : ''}"
              onclick="window.updateDayType(${idx}, '${m.key}')">${m.label}</button>
    `).join('');

    let bodyHtml = '';

    if (isLeave) {
      const leaveTitle = day.type === 'SICK_LEAVE' ? 'Sick Leave' : 'Casual Leave';
      bodyHtml = `
        <div class="leave-banner">
          <h3>${leaveTitle} Logged for ${meta.dayName}</h3>
          <p>Generates 1 Non-Production row with 0 quantity and 0 minutes. No audit, break, or documentation rows will be added.</p>
        </div>
      `;
    } else {
      const pct = (mins) => ((mins / 540) * 100).toFixed(1);
      const basePool = isHalfDay ? 210 : 420;

      bodyHtml = `
        <div class="workspace-body">
          <!-- Left: Audit Counts & 540m Budget -->
          <div class="panel-box">
            <div class="panel-title">
              <span>QA Call Audit Counts</span>
              <span class="pill ${alloc.isBalanced ? 'pill-normal' : 'pill-warn'}">
                ${alloc.isBalanced ? '✓ Balanced 540 / 540m' : `⚠️ ${alloc.totalDayTime} / 540m`}
              </span>
            </div>

            <div class="audit-grid">
              <div class="audit-box">
                <div class="audit-box-label">
                  <span>Voice Audits</span>
                  <span style="color: var(--kinpaku-gold);">●</span>
                </div>
                <div class="audit-input-row">
                  <input type="number" min="0" max="300" class="audit-number-input"
                         value="${day.voiceCount}"
                         oninput="window.updateDayInput(${idx}, 'voiceCount', this.value)">
                </div>
                <div id="voiceMinReadout" class="audit-minutes-readout">
                  → ${alloc.voiceTime} min allocated
                </div>
              </div>

              <div class="audit-box">
                <div class="audit-box-label">
                  <span>Non-Voice Audits</span>
                  <span style="color: var(--patina-teal);">●</span>
                </div>
                <div class="audit-input-row">
                  <input type="number" min="0" max="300" class="audit-number-input"
                         value="${day.nonVoiceCount}"
                         oninput="window.updateDayInput(${idx}, 'nonVoiceCount', this.value)">
                </div>
                <div id="nonVoiceMinReadout" class="audit-minutes-readout">
                  → ${alloc.nonVoiceTime} min allocated
                </div>
              </div>
            </div>

            <div class="budget-section" id="budgetSection">
              <div class="budget-header">
                <span>Audit Pool: <strong>${alloc.availableAuditPool}m</strong> (${basePool}m − ${alloc.sessionMinutes}m mtg)</span>
                <span>Total: <strong>${alloc.totalDayTime}m</strong> / 540m</span>
              </div>
              <div class="budget-bar">
                <div class="bar-seg bar-voice" style="width: ${pct(alloc.voiceTime)}%" title="Voice: ${alloc.voiceTime}m"></div>
                <div class="bar-seg bar-nonvoice" style="width: ${pct(alloc.nonVoiceTime)}%" title="Non-Voice: ${alloc.nonVoiceTime}m"></div>
                <div class="bar-seg bar-session" style="width: ${pct(alloc.sessionMinutes)}%" title="Meetings/Calibrations: ${alloc.sessionMinutes}m"></div>
                <div class="bar-seg bar-leave" style="width: ${pct(alloc.leaveTime)}%" title="Half Day Leave: ${alloc.leaveTime}m"></div>
                <div class="bar-seg bar-break" style="width: ${pct(alloc.breakTime + alloc.docTime)}%" title="Break + Doc: ${alloc.breakTime + alloc.docTime}m"></div>
              </div>
              <div class="budget-legend">
                <span class="legend-item"><span class="legend-dot" style="background: var(--kinpaku-gold)"></span>Voice (${alloc.voiceTime}m)</span>
                <span class="legend-item"><span class="legend-dot" style="background: var(--patina-teal)"></span>Non-Voice (${alloc.nonVoiceTime}m)</span>
                ${alloc.sessionMinutes > 0 ? `<span class="legend-item"><span class="legend-dot" style="background: var(--session-amber)"></span>Sessions (${alloc.sessionMinutes}m)</span>` : ''}
                ${alloc.leaveTime > 0 ? `<span class="legend-item"><span class="legend-dot" style="background: #6366f1"></span>Half Leave (${alloc.leaveTime}m)</span>` : ''}
                <span class="legend-item"><span class="legend-dot" style="background: #4b5563"></span>Break+Doc (${alloc.breakTime + alloc.docTime}m)</span>
              </div>
            </div>
          </div>

          <!-- Right: Meetings, Calibrations & Sessions -->
          <div class="panel-box">
            <div class="panel-title">
              <span>Meetings &amp; Calibrations</span>
              <span style="color: var(--session-amber); font-variant-numeric: tabular-nums;">−${alloc.sessionMinutes}m from audit</span>
            </div>

            <div class="session-toolbar">
              <button type="button" class="btn btn-secondary btn-sm" onclick="window.addPresetSession(${idx}, 'CALIBRATION')">
                + Calibration (60m)
              </button>
              <button type="button" class="btn btn-secondary btn-sm" onclick="window.addPresetSession(${idx}, 'MEETING')">
                + Meeting (60m)
              </button>
              <button type="button" class="btn btn-secondary btn-sm" onclick="window.addCustomSession(${idx})">
                + Custom Session
              </button>
            </div>

            <div class="session-cards">
              ${day.sessions.length === 0 ? `
                <div class="empty-sessions">
                  No meetings or calibrations on ${meta.dayName}. Full ${basePool} minutes allocated to QA Call Audits.
                </div>
              ` : day.sessions.map((sess, sIdx) => `
                <div class="session-card-inline">
                  <div class="session-card-top">
                    <select onchange="window.updateSessionField(${idx}, ${sIdx}, 'type', this.value)">
                      <option value="Calibration" ${sess.level1 === 'Calibration' ? 'selected' : ''}>Calibration</option>
                      <option value="Meeting" ${sess.level1 === 'Meeting' ? 'selected' : ''}>Meeting</option>
                      <option value="Session" ${sess.level1 === 'Session' ? 'selected' : ''}>Session</option>
                    </select>
                    <input type="text" value="${escapeAttr(sess.level2)}" placeholder="Session Title (Level-2)"
                           oninput="window.updateSessionField(${idx}, ${sIdx}, 'level2', this.value)">
                    <input type="number" min="5" max="420" step="5" value="${sess.time}" title="Duration in minutes"
                           oninput="window.updateSessionField(${idx}, ${sIdx}, 'time', this.value)">
                    <button type="button" class="btn-icon-remove" title="Remove session"
                            onclick="window.removeSession(${idx}, ${sIdx})">&times;</button>
                  </div>
                  <input type="text" class="session-notes-input" value="${escapeAttr(sess.level3)}"
                         placeholder="Topic / Agenda Notes (Level-3)"
                         oninput="window.updateSessionField(${idx}, ${sIdx}, 'level3', this.value)">
                </div>
              `).join('')}
            </div>
          </div>
        </div>
      `;
    }

    elWorkspace.innerHTML = `
      <div class="workspace-header">
        <div class="workspace-title-group">
          <h2>
            ${meta.dayName} — ${formatDateShort(meta.completionDate)}
          </h2>
          <div class="workspace-dates">
            <span>Receiving Date (-3d): <strong>${formatDateShort(meta.receivingDate)}</strong></span>
            <span>Completion &amp; Cut-off: <strong>${formatDateShort(meta.completionDate)}</strong></span>
          </div>
        </div>

        <div class="segmented-control" role="group" aria-label="Day Work Mode">
          ${modeButtons}
        </div>
      </div>

      ${bodyHtml}
    `;
  }

  function escapeAttr(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/"/g, '&quot;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function renderPreviewTable() {
    if (elPreviewTableHead.children.length === 0) {
      elPreviewTableHead.innerHTML = APP_CONSTANTS.HEADERS.map(h => `<th>${h}</th>`).join('');
    }

    const rows = generateReportRows(state.days, state.profile);
    elRowCountBadge.textContent = rows.length;

    elPreviewTableBody.innerHTML = rows.map(row => {
      return `<tr>${row.map((val, colIdx) => {
        if ([10, 11, 12].includes(colIdx) && val instanceof Date) {
          return `<td>${formatDateShort(val)}</td>`;
        }
        return `<td>${val !== null && val !== undefined ? val : ''}</td>`;
      }).join('')}</tr>`;
    }).join('');
  }

  function renderPreviewVisibility() {
    if (state.previewOpen) {
      elPreviewBody.classList.add('open');
      elPreviewChevron.innerHTML = 'Hide Preview &#9652;';
    } else {
      elPreviewBody.classList.remove('open');
      elPreviewChevron.innerHTML = 'Show Preview &#9662;';
    }
  }

  function refreshAll() {
    renderWeekRange();
    renderDayStrip();
    renderActiveWorkspace();
    renderPreviewTable();
    renderPreviewVisibility();
  }

  // Soft-update without re-rendering inputs so typing never loses focus
  function refreshCalculationsOnly() {
    const day = state.days[state.activeDayIndex];
    const alloc = calculateDayAllocation(day);
    const vEl = document.getElementById('voiceMinReadout');
    const nvEl = document.getElementById('nonVoiceMinReadout');
    if (vEl) vEl.textContent = `→ ${alloc.voiceTime} min allocated`;
    if (nvEl) nvEl.textContent = `→ ${alloc.nonVoiceTime} min allocated`;

    const budgetEl = document.getElementById('budgetSection');
    if (budgetEl) {
      const isHalfDay = day.type === 'HALF_DAY_LEAVE';
      const basePool = isHalfDay ? 210 : 420;
      const pct = (mins) => ((mins / 540) * 100).toFixed(1);
      budgetEl.innerHTML = `
        <div class="budget-header">
          <span>Audit Pool: <strong>${alloc.availableAuditPool}m</strong> (${basePool}m − ${alloc.sessionMinutes}m mtg)</span>
          <span>Total: <strong>${alloc.totalDayTime}m</strong> / 540m</span>
        </div>
        <div class="budget-bar">
          <div class="bar-seg bar-voice" style="width: ${pct(alloc.voiceTime)}%"></div>
          <div class="bar-seg bar-nonvoice" style="width: ${pct(alloc.nonVoiceTime)}%"></div>
          <div class="bar-seg bar-session" style="width: ${pct(alloc.sessionMinutes)}%"></div>
          <div class="bar-seg bar-leave" style="width: ${pct(alloc.leaveTime)}%"></div>
          <div class="bar-seg bar-break" style="width: ${pct(alloc.breakTime + alloc.docTime)}%"></div>
        </div>
        <div class="budget-legend">
          <span class="legend-item"><span class="legend-dot" style="background: var(--kinpaku-gold)"></span>Voice (${alloc.voiceTime}m)</span>
          <span class="legend-item"><span class="legend-dot" style="background: var(--patina-teal)"></span>Non-Voice (${alloc.nonVoiceTime}m)</span>
          ${alloc.sessionMinutes > 0 ? `<span class="legend-item"><span class="legend-dot" style="background: var(--session-amber)"></span>Sessions (${alloc.sessionMinutes}m)</span>` : ''}
          ${alloc.leaveTime > 0 ? `<span class="legend-item"><span class="legend-dot" style="background: #6366f1"></span>Half Leave (${alloc.leaveTime}m)</span>` : ''}
          <span class="legend-item"><span class="legend-dot" style="background: #4b5563"></span>Break+Doc (${alloc.breakTime + alloc.docTime}m)</span>
        </div>
      `;
    }

    renderDayStrip();
    renderPreviewTable();
  }

  // Global Actions
  root.selectDay = function (idx) {
    state.activeDayIndex = idx;
    saveState();
    refreshAll();
  };

  root.updateDayType = function (dayIndex, newType) {
    state.days[dayIndex].type = newType;
    saveState();
    refreshAll();
  };

  root.updateDayInput = function (dayIndex, field, value) {
    state.days[dayIndex][field] = Math.max(0, parseInt(value, 10) || 0);
    saveState();
    refreshCalculationsOnly();
  };

  root.addPresetSession = function (dayIndex, presetKey) {
    const preset = SESSION_PLACEHOLDERS[presetKey];
    if (!preset) return;
    state.days[dayIndex].sessions.push({ ...preset, id: Date.now() });
    saveState();
    refreshAll();
  };

  root.addCustomSession = function (dayIndex) {
    state.days[dayIndex].sessions.push({
      source: 'Meeting',
      frequency: 'Ad hoc',
      production: 'Production',
      sla: 'Yes',
      level1: 'Meeting',
      level2: 'Team Sync Session',
      level3: 'Discussion & QA Alignment',
      quantity: 1,
      time: 60,
      id: Date.now()
    });
    saveState();
    refreshAll();
  };

  root.removeSession = function (dayIndex, sessionIndex) {
    state.days[dayIndex].sessions.splice(sessionIndex, 1);
    saveState();
    refreshAll();
  };

  root.updateSessionField = function (dayIndex, sessionIndex, field, value) {
    const sess = state.days[dayIndex].sessions[sessionIndex];
    if (!sess) return;
    if (field === 'type') {
      sess.source = value;
      sess.level1 = value;
      saveState();
      refreshCalculationsOnly();
    } else if (field === 'time') {
      sess.time = Math.max(0, parseInt(value, 10) || 0);
      saveState();
      refreshCalculationsOnly();
    } else {
      sess[field] = value;
      saveState();
      renderPreviewTable();
    }
  };

  // Header & Config Listeners
  elProcessorName.value = state.profile.name;
  elProcessorName.addEventListener('input', (e) => {
    state.profile.name = e.target.value;
    saveState();
    renderPreviewTable();
  });

  elQcBy.value = state.profile.qcBy;
  elQcBy.addEventListener('input', (e) => {
    state.profile.qcBy = e.target.value;
    saveState();
    renderPreviewTable();
  });

  elStartDate.value = state.startDate;
  function updateStartDate(newDateStr) {
    state.startDate = newDateStr;
    elStartDate.value = newDateStr;
    const dates = resolveWeekDates(newDateStr);
    state.days.forEach((day, idx) => {
      day.dateMeta = dates[idx];
    });
    saveState();
    refreshAll();
  }

  elStartDate.addEventListener('change', (e) => {
    if (e.target.value) updateStartDate(e.target.value);
  });

  function shiftWeek(offsetDays) {
    const [y, m, d] = state.startDate.split('-').map(Number);
    const shifted = new Date(y, m - 1, d + offsetDays, 12, 0, 0);
    const pad = (n) => String(n).padStart(2, '0');
    updateStartDate(`${shifted.getFullYear()}-${pad(shifted.getMonth() + 1)}-${pad(shifted.getDate())}`);
  }

  elBtnPrevWeek.addEventListener('click', () => shiftWeek(-7));
  elBtnNextWeek.addEventListener('click', () => shiftWeek(7));

  elBtnResetWeek.addEventListener('click', () => {
    if (!confirm('Reset all 5 days of this week back to default values?')) return;
    state.days = createDefaultDays(state.startDate);
    state.activeDayIndex = 0;
    saveState();
    refreshAll();
  });

  elPreviewToggleBar.addEventListener('click', () => {
    state.previewOpen = !state.previewOpen;
    saveState();
    renderPreviewVisibility();
  });

  elBtnDemoShuvo.addEventListener('click', () => {
    updateStartDate('2026-10-25');
    state.profile.name = 'Jahidul Islam Shuvo';
    elProcessorName.value = state.profile.name;

    state.days[0].type = 'HALF_DAY_LEAVE';
    state.days[0].voiceCount = 0;
    state.days[0].nonVoiceCount = 30;
    state.days[0].sessions = [];

    state.days[1].type = 'NORMAL';
    state.days[1].voiceCount = 35;
    state.days[1].nonVoiceCount = 0;
    state.days[1].sessions = [];

    state.days[2].type = 'NORMAL';
    state.days[2].voiceCount = 60;
    state.days[2].nonVoiceCount = 0;
    state.days[2].sessions = [];

    state.days[3].type = 'NORMAL';
    state.days[3].voiceCount = 60;
    state.days[3].nonVoiceCount = 0;
    state.days[3].sessions = [];

    state.days[4].type = 'NORMAL';
    state.days[4].voiceCount = 48;
    state.days[4].nonVoiceCount = 12;
    state.days[4].sessions = [];

    state.activeDayIndex = 0;
    saveState();
    refreshAll();
  });

  elBtnDemoRony.addEventListener('click', () => {
    updateStartDate('2026-11-08');
    state.profile.name = 'Intakhab Didar Rony';
    elProcessorName.value = state.profile.name;

    state.days[0].type = 'NORMAL';
    state.days[0].voiceCount = 51;
    state.days[0].nonVoiceCount = 0;
    state.days[0].sessions = [];

    state.days[1].type = 'NORMAL';
    state.days[1].voiceCount = 31;
    state.days[1].nonVoiceCount = 30;
    state.days[1].sessions = [{
      source: 'Calibration',
      frequency: 'Ad hoc',
      production: 'Production',
      sla: 'Yes',
      level1: 'Calibration',
      level2: 'Calibration with Mia Bhai',
      level3: 'Calibration about fail calls',
      quantity: 1,
      time: 60,
      id: 1
    }];

    state.days[2].type = 'NORMAL';
    state.days[2].voiceCount = 26;
    state.days[2].nonVoiceCount = 35;
    state.days[2].sessions = [];

    state.days[3].type = 'SICK_LEAVE';
    state.days[3].voiceCount = 0;
    state.days[3].nonVoiceCount = 0;
    state.days[3].sessions = [];

    state.days[4].type = 'NORMAL';
    state.days[4].voiceCount = 46;
    state.days[4].nonVoiceCount = 14;
    state.days[4].sessions = [{
      source: 'Meeting',
      frequency: 'Ad hoc',
      production: 'Production',
      sla: 'Yes',
      level1: 'Meeting',
      level2: 'Meeting with GP & Genex',
      level3: 'Meeting about Complaint Management',
      quantity: 1,
      time: 60,
      id: 2
    }];

    state.activeDayIndex = 1;
    saveState();
    refreshAll();
  });

  elBtnImportAudit.addEventListener('click', () => {
    elAuditFileInput.value = '';
    elAuditFileInput.click();
  });

  elAuditFileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const result = await parseAuditWorkbook(buffer);

      // Align week Sunday to the imported file's week
      state.startDate = result.weekSunday;
      elStartDate.value = result.weekSunday;
      const dates = resolveWeekDates(result.weekSunday);

      state.days.forEach((day, idx) => {
        day.dateMeta = dates[idx];
        const counts = result.dailyCounts[day.dateMeta.isoDate];
        if (counts) {
          day.voiceCount = counts.voice;
          day.nonVoiceCount = counts.nonVoice;
          if (day.type === 'SICK_LEAVE' || day.type === 'CASUAL_LEAVE') {
            day.type = 'NORMAL';
          }
        }
      });

      state.activeDayIndex = 0;
      saveState();
      refreshAll();

      const indicator = document.getElementById('saveIndicator');
      if (indicator) {
        indicator.textContent = `✓ Imported ${result.sortedDates.length} days (${result.totalVoice}V / ${result.totalNonVoice}NV)`;
      }
    } catch (err) {
      alert('Failed to import Audit file: ' + err.message);
    }
  });

  function showSyncFeedback(type, message) {
    elSyncFeedback.className = `sync-feedback show ${type}`;
    elSyncFeedbackText.textContent = message;
  }

  elBtnDismissFeedback.addEventListener('click', () => {
    elSyncFeedback.className = 'sync-feedback';
  });

  elBtnSyncSheet.addEventListener('click', async () => {
    const originalText = elBtnSyncSheet.textContent;
    try {
      elBtnSyncSheet.disabled = true;
      elBtnSyncSheet.textContent = 'Syncing...';
      elSyncFeedback.className = 'sync-feedback';

      const result = await syncFromGoogleSheet(DEFAULT_SHEET_URL, elProcessorName.value);

      // Update autocomplete suggestions from live sheet
      if (elNameSuggestions && result.availableNames.length > 0) {
        elNameSuggestions.innerHTML = result.availableNames
          .map((n) => `<option value="${escapeAttr(n)}"></option>`)
          .join('');
      }

      // Normalize processor name to exact sheet spelling
      state.profile.name = result.matchedName;
      elProcessorName.value = result.matchedName;

      // Match the currently selected Sun-Thu week dates
      let matchedDays = 0;
      let totalV = 0;
      let totalNV = 0;

      state.days.forEach((day) => {
        const counts = result.dailyCounts[day.dateMeta.isoDate];
        if (counts) {
          day.voiceCount = counts.voice;
          day.nonVoiceCount = counts.nonVoice;
          totalV += counts.voice;
          totalNV += counts.nonVoice;
          matchedDays++;
          if ((counts.voice > 0 || counts.nonVoice > 0) && (day.type === 'SICK_LEAVE' || day.type === 'CASUAL_LEAVE')) {
            day.type = 'NORMAL';
          }
        }
      });

      if (matchedDays === 0) {
        const firstIso = state.days[0].dateMeta.isoDate;
        const lastIso = state.days[4].dateMeta.isoDate;
        showSyncFeedback(
          'error',
          `Found "${result.matchedName}", but no audit rows matched selected week (${firstIso} to ${lastIso}). Try changing the Week Starting date.`
        );
        return;
      }

      saveState();
      refreshAll();
      showSyncFeedback(
        'success',
        `✓ Synced ${matchedDays} days from Google Sheet for ${result.matchedName} (${totalV} Voice, ${totalNV} Non-Voice).`
      );
    } catch (err) {
      showSyncFeedback('error', err.message);
    } finally {
      elBtnSyncSheet.disabled = false;
      elBtnSyncSheet.textContent = originalText;
    }
  });

  elBtnExport.addEventListener('click', async () => {
    const originalHtml = elBtnExport.innerHTML;
    try {
      elBtnExport.disabled = true;
      elBtnExport.textContent = 'Generating Excel...';
      await exportReportToExcel(state.days, state.profile);
    } catch (err) {
      alert('Export failed: ' + err.message);
    } finally {
      setTimeout(() => {
        elBtnExport.disabled = false;
        elBtnExport.innerHTML = originalHtml;
      }, 600);
    }
  });

  // Initial Render
  saveState();
  refreshAll();
})(window);
