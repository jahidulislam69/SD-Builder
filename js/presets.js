// Presets, default constants, and canonical taxonomy for SD Daily Report
(function (root) {
  const SD = (root.SD = root.SD || {});

  SD.APP_CONSTANTS = {
    UNIT: 'QA Executive',
    QC_BY_DEFAULT: 'Md Helal Hossain',
    CATEGORY: 'Office Work',
    SLA_DEFAULT: 'Yes',

    // Exact column headers matching historical SD reports
    HEADERS: [
      'SL no',
      'Soruce', // Intentional historical spelling in original template
      'Processor name',
      'Frequency',
      'Unit',
      'Production/Non Production',
      'SLA',
      'Level-1',
      'Level-2',
      'Level-3',
      'Complete request receiving date(MM/DD/YY)',
      ' Completion Date (MM/DD/YY)', // Intentional leading space
      ' Cut off/Completion Date (MM/DD/YY)', // Intentional leading space
      'Quantity/No of request',
      'Time taken for delivery',
      'QC By',
      'Category\u2019 ', // Exact unicode character matching original files
      'Month',
      'Year'
    ],

    COLUMN_WIDTHS: [
      13.0, 16.0, 22.0, 12.0, 14.0, 26.0, 8.0, 30.0, 22.0, 44.0,
      30.0, 28.0, 28.0, 22.0, 22.0, 18.0, 14.0, 14.0, 12.0
    ]
  };

  // Activity taxonomy presets
  SD.ACTIVITY_PRESETS = {
    VOICE_AUDIT: {
      source: 'VDI log In',
      frequency: 'Daily',
      production: 'Production',
      sla: 'Yes',
      level1: 'QA Call Audit',
      level2: 'Voice',
      level3: 'Audit & Pass/Fail submit in Protal'
    },
    NON_VOICE_AUDIT: {
      source: 'VDI log In',
      frequency: 'Daily',
      production: 'Production',
      sla: 'Yes',
      level1: 'QA Call Audit',
      level2: 'Non Voice',
      level3: 'Audit & Pass/Fail submit in Protal'
    },
    BREAK: {
      source: 'Break',
      frequency: 'Daily',
      production: 'Production',
      sla: 'Yes',
      level1: 'Lunch, Prayer and Other Break',
      level2: 'Lunch, Prayer and Other Break',
      level3: 'Lunch, Prayer and Other Break',
      quantity: 1,
      time: 60
    },
    DOCUMENTATION: {
      source: 'VDI log In',
      frequency: 'Daily',
      production: 'Production',
      sla: 'Yes',
      level1: 'Documentation',
      level2: 'Keep Daily Log',
      level3: 'Audit report & SD report update',
      quantity: 2,
      time: 60
    },
    HALF_DAY_LEAVE: {
      source: 'Half Day',
      frequency: 'Ad hoc',
      production: 'Non Production',
      sla: 'Yes',
      level1: 'Half Day Leave',
      level2: 'Half Day Leave',
      level3: 'Half Day Leave',
      quantity: 1,
      time: 270
    },
    SICK_LEAVE: {
      source: 'Leave',
      frequency: 'Ad hoc',
      production: 'Non Production',
      sla: 'Yes',
      level1: 'Sick Leave',
      level2: 'Sick Leave',
      level3: 'Sick Leave',
      quantity: 0,
      time: 0
    },
    CASUAL_LEAVE: {
      source: 'Leave',
      frequency: 'Daily',
      production: 'Non Production',
      sla: 'Yes',
      level1: 'Casual Leave',
      level2: 'Casual Leave',
      level3: 'Casual Leave',
      quantity: 0,
      time: 0
    }
  };

  // Demo placeholders for Ad hoc Sessions / Meetings / Calibrations
  SD.SESSION_PLACEHOLDERS = {
    CALIBRATION: {
      source: 'Calibration',
      frequency: 'Ad hoc',
      production: 'Production',
      sla: 'Yes',
      level1: 'Calibration',
      level2: 'Calibration with Mia Bhai',
      level3: 'Calibration about fail calls',
      quantity: 1,
      time: 60
    },
    MEETING: {
      source: 'Meeting',
      frequency: 'Ad hoc',
      production: 'Production',
      sla: 'Yes',
      level1: 'Meeting',
      level2: 'Meeting with GP & Genex',
      level3: 'Meeting about Complaint Management',
      quantity: 1,
      time: 60
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : window);
