// Core calculation and date math engine
(function (root) {
  const SD = (root.SD = root.SD || {});

  const MONTH_NAMES = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  /**
   * Given a Sunday start date string (YYYY-MM-DD), resolves Sunday through Thursday
   * and calculates the 3-day previous receiving date.
   */
  SD.resolveWeekDates = function (sundayStr) {
    const [y, m, d] = sundayStr.split('-').map(Number);
    // Use noon to avoid timezone / daylight saving day shifts
    const sunday = new Date(y, m - 1, d, 12, 0, 0);

    const days = [];
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday'];

    for (let i = 0; i < 5; i++) {
      const compDate = new Date(sunday);
      compDate.setDate(sunday.getDate() + i);

      // Complete request receiving date is 3 calendar days before completion date
      const recDate = new Date(compDate);
      recDate.setDate(compDate.getDate() - 3);

      const pad = (n) => String(n).padStart(2, '0');
      const compIso = `${compDate.getFullYear()}-${pad(compDate.getMonth() + 1)}-${pad(compDate.getDate())}`;

      days.push({
        index: i,
        dayName: dayNames[i],
        completionDate: compDate,
        cutoffDate: new Date(compDate),
        receivingDate: recDate,
        month: MONTH_NAMES[compDate.getMonth()],
        year: compDate.getFullYear(),
        isoDate: compIso
      });
    }

    return days;
  };

  /**
   * Calculates time allocation for a single day based on day type, counts, and sessions.
   */
  SD.calculateDayAllocation = function (day) {
    const { type, voiceCount, nonVoiceCount, sessions = [] } = day;

    if (type === 'SICK_LEAVE' || type === 'CASUAL_LEAVE') {
      return {
        type,
        baseAuditPool: 0,
        sessionMinutes: 0,
        availableAuditPool: 0,
        voiceTime: 0,
        nonVoiceTime: 0,
        breakTime: 0,
        docTime: 0,
        leaveTime: 0,
        totalDayTime: 0,
        isBalanced: true
      };
    }

    const sessionMinutes = sessions.reduce((acc, s) => acc + (Number(s.time) || 0), 0);

    const isHalfDay = (type === 'HALF_DAY_LEAVE');
    const baseAuditPool = isHalfDay ? 210 : 420;
    const breakTime = isHalfDay ? 30 : 60;
    const docTime = isHalfDay ? 30 : 60;
    const leaveTime = isHalfDay ? 270 : 0;

    // Meeting / calibration time is deducted from the QA Call Audit pool
    const availableAuditPool = Math.max(0, baseAuditPool - sessionMinutes);

    const vCount = Number(voiceCount) || 0;
    const nvCount = Number(nonVoiceCount) || 0;
    const totalCount = vCount + nvCount;

    let voiceTime = 0;
    let nonVoiceTime = 0;

    if (totalCount > 0) {
      if (vCount > 0 && nvCount === 0) {
        voiceTime = availableAuditPool;
      } else if (vCount === 0 && nvCount > 0) {
        nonVoiceTime = availableAuditPool;
      } else {
        voiceTime = Math.round((vCount / totalCount) * availableAuditPool);
        nonVoiceTime = availableAuditPool - voiceTime;
      }
    } else if (availableAuditPool > 0) {
      voiceTime = availableAuditPool;
    }

    const totalDayTime = voiceTime + nonVoiceTime + sessionMinutes + breakTime + docTime + leaveTime;
    const isBalanced = (totalDayTime === 540);

    return {
      type,
      baseAuditPool,
      sessionMinutes,
      availableAuditPool,
      voiceTime,
      nonVoiceTime,
      breakTime,
      docTime,
      leaveTime,
      totalDayTime,
      isBalanced
    };
  };

  /**
   * Builds the canonical list of 19-column spreadsheet rows from the week data.
   */
  SD.generateReportRows = function (weekData, profile) {
    const { APP_CONSTANTS, ACTIVITY_PRESETS } = SD;
    const rows = [];
    let slNo = 1;

    for (const day of weekData) {
      const alloc = SD.calculateDayAllocation(day);
      const { completionDate, cutoffDate, receivingDate, month, year } = day.dateMeta;

      if (day.type === 'SICK_LEAVE' || day.type === 'CASUAL_LEAVE') {
        const preset = ACTIVITY_PRESETS[day.type];
        rows.push(createRow(slNo++, {
          source: preset.source,
          processorName: profile.name,
          frequency: preset.frequency,
          unit: profile.unit || APP_CONSTANTS.UNIT,
          production: preset.production,
          sla: preset.sla,
          level1: preset.level1,
          level2: preset.level2,
          level3: preset.level3,
          receivingDate: completionDate,
          completionDate,
          cutoffDate,
          quantity: 0,
          time: 0,
          qcBy: profile.qcBy || APP_CONSTANTS.QC_BY_DEFAULT,
          category: APP_CONSTANTS.CATEGORY,
          month,
          year
        }));
        continue;
      }

      if (day.type === 'HALF_DAY_LEAVE') {
        const preset = ACTIVITY_PRESETS.HALF_DAY_LEAVE;
        rows.push(createRow(slNo++, {
          source: preset.source,
          processorName: profile.name,
          frequency: preset.frequency,
          unit: profile.unit || APP_CONSTANTS.UNIT,
          production: preset.production,
          sla: preset.sla,
          level1: preset.level1,
          level2: preset.level2,
          level3: preset.level3,
          receivingDate: completionDate,
          completionDate,
          cutoffDate,
          quantity: preset.quantity,
          time: preset.time,
          qcBy: profile.qcBy || APP_CONSTANTS.QC_BY_DEFAULT,
          category: APP_CONSTANTS.CATEGORY,
          month,
          year
        }));
      }

      const voicePreset = ACTIVITY_PRESETS.VOICE_AUDIT;
      rows.push(createRow(slNo++, {
        source: voicePreset.source,
        processorName: profile.name,
        frequency: voicePreset.frequency,
        unit: profile.unit || APP_CONSTANTS.UNIT,
        production: voicePreset.production,
        sla: voicePreset.sla,
        level1: voicePreset.level1,
        level2: voicePreset.level2,
        level3: voicePreset.level3,
        receivingDate,
        completionDate,
        cutoffDate,
        quantity: Number(day.voiceCount) || 0,
        time: alloc.voiceTime,
        qcBy: profile.qcBy || APP_CONSTANTS.QC_BY_DEFAULT,
        category: APP_CONSTANTS.CATEGORY,
        month,
        year
      }));

      const nonVoicePreset = ACTIVITY_PRESETS.NON_VOICE_AUDIT;
      rows.push(createRow(slNo++, {
        source: nonVoicePreset.source,
        processorName: profile.name,
        frequency: nonVoicePreset.frequency,
        unit: profile.unit || APP_CONSTANTS.UNIT,
        production: nonVoicePreset.production,
        sla: nonVoicePreset.sla,
        level1: nonVoicePreset.level1,
        level2: nonVoicePreset.level2,
        level3: nonVoicePreset.level3,
        receivingDate,
        completionDate,
        cutoffDate,
        quantity: Number(day.nonVoiceCount) || 0,
        time: alloc.nonVoiceTime,
        qcBy: profile.qcBy || APP_CONSTANTS.QC_BY_DEFAULT,
        category: APP_CONSTANTS.CATEGORY,
        month,
        year
      }));

      const breakPreset = ACTIVITY_PRESETS.BREAK;
      rows.push(createRow(slNo++, {
        source: breakPreset.source,
        processorName: profile.name,
        frequency: breakPreset.frequency,
        unit: profile.unit || APP_CONSTANTS.UNIT,
        production: breakPreset.production,
        sla: breakPreset.sla,
        level1: breakPreset.level1,
        level2: breakPreset.level2,
        level3: breakPreset.level3,
        receivingDate: completionDate,
        completionDate,
        cutoffDate,
        quantity: 1,
        time: alloc.breakTime,
        qcBy: profile.qcBy || APP_CONSTANTS.QC_BY_DEFAULT,
        category: APP_CONSTANTS.CATEGORY,
        month,
        year
      }));

      if (day.sessions && day.sessions.length > 0) {
        for (const session of day.sessions) {
          rows.push(createRow(slNo++, {
            source: session.source || 'Meeting',
            processorName: profile.name,
            frequency: session.frequency || 'Ad hoc',
            unit: profile.unit || APP_CONSTANTS.UNIT,
            production: session.production || 'Production',
            sla: session.sla || 'Yes',
            level1: session.level1 || 'Meeting',
            level2: session.level2 || 'General Session',
            level3: session.level3 || 'Discussion & Follow-up',
            receivingDate: completionDate,
            completionDate,
            cutoffDate,
            quantity: Number(session.quantity) || 1,
            time: Number(session.time) || 0,
            qcBy: profile.qcBy || APP_CONSTANTS.QC_BY_DEFAULT,
            category: APP_CONSTANTS.CATEGORY,
            month,
            year
          }));
        }
      }

      const docPreset = ACTIVITY_PRESETS.DOCUMENTATION;
      rows.push(createRow(slNo++, {
        source: docPreset.source,
        processorName: profile.name,
        frequency: docPreset.frequency,
        unit: profile.unit || APP_CONSTANTS.UNIT,
        production: docPreset.production,
        sla: docPreset.sla,
        level1: docPreset.level1,
        level2: docPreset.level2,
        level3: docPreset.level3,
        receivingDate: completionDate,
        completionDate,
        cutoffDate,
        quantity: 2,
        time: alloc.docTime,
        qcBy: profile.qcBy || APP_CONSTANTS.QC_BY_DEFAULT,
        category: APP_CONSTANTS.CATEGORY,
        month,
        year
      }));
    }

    return rows;
  };

  function createRow(sl, fields) {
    return [
      sl,
      fields.source,
      fields.processorName,
      fields.frequency,
      fields.unit,
      fields.production,
      fields.sla,
      fields.level1,
      fields.level2,
      fields.level3,
      fields.receivingDate,
      fields.completionDate,
      fields.cutoffDate,
      fields.quantity,
      fields.time,
      fields.qcBy,
      fields.category,
      fields.month,
      fields.year
    ];
  }
})(typeof globalThis !== 'undefined' ? globalThis : window);
