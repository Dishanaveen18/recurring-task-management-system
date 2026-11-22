function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Task Operations')
    .addItem('Initialize / Refresh Tasks', 'initializeSystem')
    .addItem('Submit Selected Task', 'submitSelectedTask')
    .addItem('Check Manager Alerts', 'sendManagerOverdueAlerts')
    .addItem('Enable Automatic Manager Alerts', 'setupManagerAlertTrigger')
    .addItem('Acknowledge Selected Alert', 'acknowledgeSelectedAlert')
    .addItem('Enable Leadership Escalation','setupLeadershipEscalationTrigger')
    .addToUi();
}




function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const master = ss.getSheetByName('Task_Master');
  const log = ss.getSheetByName('Task_Cycles_Log');

  if (!master || !log) {
    throw new Error('Check that Task_Master and Task_Cycles_Log exist.');
  }

  const tasks = master.getDataRange().getValues();
  const lastLogRow = log.getLastRow();

  const existingCycles = lastLogRow > 1
    ? log.getRange(2, 1, lastLogRow - 1, 1)
        .getValues().flat()
    : [];

  const known = new Set(existingCycles);
  const newCycles = [];

  for (let i = 1; i < tasks.length; i++) {
    const t = tasks[i];

    const taskId = t[0];
    const taskName = t[1];
    const frequency = Number(t[4]);
    const firstDueDate = t[5];
    const dueTime = t[6];
    const active = t[8] === true ||
      String(t[8]).toUpperCase() === 'TRUE';

    if (!taskId || !active) continue;

    if (!Number.isFinite(frequency) || frequency < 1 ||
        !(firstDueDate instanceof Date) ||
        isNaN(firstDueDate.getTime())) {
      throw new Error(
        'Check Frequency_Days and First_Due_Date for ' + taskId
      );
    }

    const cycleId = taskId + '-C1';

    // Do not create the same initial cycle twice.
    if (known.has(cycleId)) continue;

    const due = new Date(firstDueDate);
    let hours = 17;
    let minutes = 0;

    if (dueTime instanceof Date) {
      hours = dueTime.getHours();
      minutes = dueTime.getMinutes();
    } else if (typeof dueTime === 'string' && dueTime.trim()) {
      const match = dueTime.match(/^(\d{1,2}):(\d{2})$/);

      if (!match) {
        throw new Error('Use HH:MM for Due_Time: ' + taskId);
      }

      hours = Number(match[1]);
      minutes = Number(match[2]);
    }

    due.setHours(hours, minutes, 0, 0);

    newCycles.push([
      cycleId,
      taskId,
      1,
      due,
      getTaskStatus(due),
      '',
      '',
      '',
      '',
      ''
    ]);

    known.add(cycleId);
  }

  if (newCycles.length > 0) {
    log.getRange(
      log.getLastRow() + 1, 1, newCycles.length, 10
    ).setValues(newCycles);
  }

  refreshEmployeeView();

  SpreadsheetApp.getUi().alert(
    'Initialization complete!\n' +
    newCycles.length + ' new task cycle(s) created.'
  );
}

function getTaskStatus(due) {
  const now = new Date();

  if (due.getTime() < now.getTime()) {
    return 'Overdue';
  }

  const today = new Date(now);
  today.setHours(0, 0, 0, 0);

  const dueDay = new Date(due);
  dueDay.setHours(0, 0, 0, 0);

  return dueDay.getTime() === today.getTime()
    ? 'Pending'
    : 'Upcoming';
}


function refreshEmployeeView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('Task_Master');
  const log = ss.getSheetByName('Task_Cycles_Log');
  const view = ss.getSheetByName('Employee_Task_View');

  if (!master || !log || !view) {
    throw new Error('Please check that all three sheets exist.');
  }

  const masterData = master.getDataRange().getValues();
  const logData = log.getDataRange().getValues();

  // Remember employee inputs before refreshing.
  const oldLastRow = view.getLastRow();
  const oldRows = oldLastRow > 1
    ? view.getRange(2, 1, oldLastRow - 1, 11).getValues()
    : [];

  const savedInputs = new Map();

  oldRows.forEach(row => {
    if (row[0]) {
      savedInputs.set(row[0], [row[6], row[7], row[8]]);
    }
  });

  // Create a lookup of active tasks.
  const taskMap = new Map();

  masterData.slice(1).forEach(task => {
    const active = task[8] === true ||
      String(task[8]).toUpperCase() === 'TRUE';

    if (task[0] && active) {
      taskMap.set(task[0], task);
    }
  });

  const output = [];
  const statusUpdates = [];

  // Build the employee view in memory.
  logData.slice(1).forEach((cycle, index) => {
    if (!cycle[0] || cycle[4] === 'Completed') return;

    const task = taskMap.get(cycle[1]);
    if (!task) return;

    const due = new Date(cycle[3]);
    const status = getTaskStatus(due);

    statusUpdates.push({
      row: index + 2,
      status: status
    });

    const input = savedInputs.get(cycle[0]) || ['', '', ''];

    output.push([
      cycle[0],
      cycle[1],
      task[1],
      task[2],
      due,
      status,
      input[0],
      input[1],
      input[2],
      '',
      ''
    ]);
  });

  // Update cycle statuses together, rather than cell by cell.
  if (statusUpdates.length > 0) {
    const firstRow = statusUpdates[0].row;
    const lastRow = statusUpdates[statusUpdates.length - 1].row;

    // Read the status column once and update it in memory.
    const statuses = log.getRange(firstRow, 5, lastRow - firstRow + 1, 1)
      .getValues();

    statusUpdates.forEach(item => {
      statuses[item.row - firstRow][0] = item.status;
    });

    log.getRange(firstRow, 5, statuses.length, 1)
      .setValues(statuses);
  }

  // Clear old employee-view rows and write the new view once.
  if (oldLastRow > 1) {
    view.getRange(2, 1, oldLastRow - 1, 11).clearContent();
  }

  if (output.length > 0) {
    view.getRange(2, 1, output.length, 11).setValues(output);
  }

  if (output.length > 0) {
  view.getRange(2, 5, output.length, 1)
    .setNumberFormat('dd-mmm-yyyy hh:mm');
}
}

function diagnoseTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    'Task_Master',
    'Employee_Task_View',
    'Task_Cycles_Log',
    'Manager_Dashboard',
    'Alert_Log'
  ];

  const results = [];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);

    if (!sheet) {
      results.push(name + ': MISSING');
    } else {
      results.push(
        name +
        ': rows=' + sheet.getLastRow() +
        ', columns=' + sheet.getLastColumn()
      );
    }
  });

  Logger.log(results.join('\n'));
}

function submitSelectedTask() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const view = ss.getSheetByName('Employee_Task_View');
  const log = ss.getSheetByName('Task_Cycles_Log');
  const master = ss.getSheetByName('Task_Master');

  if (ss.getActiveSheet().getName() !== 'Employee_Task_View') {
    ui.alert('Open Employee_Task_View and select a task row first.');
    return;
  }

  const row = view.getActiveRange().getRow();

  if (row < 2) {
    ui.alert('Select a task row below the header.');
    return;
  }

  const data = view.getRange(row, 1, 1, 11).getValues()[0];
  const cycleId = data[0];
  const taskId = data[1];

  if (!cycleId || !taskId) {
    ui.alert('This row does not contain a task.');
    return;
  }

  const ticketId = String(data[6] || '').trim();
  const proof = String(data[7] || '').trim();

  const logData = log.getDataRange().getValues();
  const logIndex = logData.findIndex(
    (r, i) => i > 0 && r[0] === cycleId
  );

  if (logIndex === -1) {
    ui.alert('Cycle not found in Task_Cycles_Log.');
    return;
  }

  const cycle = logData[logIndex];

  if (cycle[4] === 'Completed') {
    ui.alert('This task cycle has already been completed.');
    return;
  }

  const masterData = master.getDataRange().getValues();
  const task = masterData.slice(1).find(r => r[0] === taskId);

  if (!task) {
    ui.alert('Task not found in Task_Master.');
    return;
  }

  const proofType = String(task[7] || '').toLowerCase();

  if (proofType.includes('ticket') && !ticketId) {
    ui.alert('Enter the Ticket_ID in column G first.');
    return;
  }

  if (proofType.includes('whatsapp') && !proof) {
    ui.alert('Enter the WhatsApp proof reference in column H first.');
    return;
  }

  const frequency = Number(task[4]);

  if (!Number.isFinite(frequency) || frequency < 1) {
    ui.alert('Check Frequency_Days in Task_Master.');
    return;
  }

  const now = new Date();
  const currentDue = new Date(cycle[3]);
  const nextDue = new Date(currentDue);
  nextDue.setDate(nextDue.getDate() + frequency);

  const cycleNumber = Number(cycle[2]) || 1;
  const nextCycleId = taskId + '-C' + (cycleNumber + 1);

  // Prevent duplicate completion and duplicate next cycles.
  const existingIds = logData.slice(1).map(r => r[0]);

  if (existingIds.includes(nextCycleId)) {
    ui.alert(
      'The next cycle already exists. No changes were made.'
    );
    return;
  }

  const email = Session.getActiveUser().getEmail();
  const completedBy = email || String(task[2] || 'Employee');

  // Mark the current cycle complete.
  log.getRange(logIndex + 1, 5, 1, 6).setValues([[
    'Completed',
    completedBy,
    now,
    ticketId,
    proof,
    nextDue
  ]]);

  // Create the next recurring cycle.
  log.appendRow([
    nextCycleId,
    taskId,
    cycleNumber + 1,
    nextDue,
    getTaskStatus(nextDue),
    '',
    '',
    '',
    '',
    ''
  ]);

  // Update the selected employee-view row.
  view.getRange(row, 6).setValue('Completed');
  view.getRange(row, 10, 1, 2).setValues([[
    completedBy,
    now
  ]]);

  // Add the next cycle to the employee view.
  view.appendRow([
    nextCycleId,
    taskId,
    task[1],
    task[2],
    nextDue,
    getTaskStatus(nextDue),
    '',
    '',
    '',
    '',
    ''
  ]);

  view.getRange(view.getLastRow(), 5)
    .setNumberFormat('dd-mmm-yyyy hh:mm');

  ui.alert(
    'Task completed successfully!\n' +
    'Completed cycle: ' + cycleId + '\n' +
    'Next cycle: ' + nextCycleId
  );
}

function sendManagerOverdueAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('Task_Cycles_Log');
  const alertLog = ss.getSheetByName('Alert_Log');

  const managerEmail = 'dishanaveen.in@gmail.com';
  const now = new Date();

  const cycles = log.getDataRange().getValues();
  const alerts = alertLog.getDataRange().getValues();

  // Remember which task cycles have already triggered an alert.
  const alertedCycles = new Set();

  alerts.slice(1).forEach(row => {
    if (row[1] && row[2] === 'Manager Overdue Alert') {
      alertedCycles.add(row[1]);
    }
  });

  let alertsSent = 0;

  cycles.slice(1).forEach((cycle, index) => {
    const cycleId = cycle[0];
    const taskId = cycle[1];
    const dueDate = cycle[3];
    const status = cycle[4];

    // Ignore empty, completed, or already-alerted cycles.
    if (!cycleId || !taskId || status === 'Completed') return;
    if (alertedCycles.has(cycleId)) return;
    if (!(dueDate instanceof Date) || isNaN(dueDate.getTime())) return;

    // Alert 15 minutes after the task deadline.
    const alertTime = new Date(dueDate.getTime() + 15 * 60 * 1000);
    if (now < alertTime) return;

    // Find the task details.
    const master = ss.getSheetByName('Task_Master');
    const tasks = master.getDataRange().getValues();
    const task = tasks.slice(1).find(row => row[0] === taskId);

    if (!task) return;

    const employee = task[2] || 'Assigned employee';
    const employeeEmail = task[3] || 'Not provided';
    const taskName = task[1] || taskId;

    const subject = 'Overdue Task Alert: ' + taskName;

    const body =
      'Hello Manager,\n\n' +
      'A recurring task has passed its deadline and is still incomplete.\n\n' +
      'Task: ' + taskName + '\n' +
      'Task ID: ' + taskId + '\n' +
      'Cycle ID: ' + cycleId + '\n' +
      'Assigned to: ' + employee + '\n' +
      'Employee email: ' + employeeEmail + '\n' +
      'Due date: ' + dueDate + '\n\n' +
      'Please review this task and acknowledge the alert in the tracker.\n\n' +
      '— Allo Innoware Task Tracker';

    // Send the email.
    GmailApp.sendEmail(managerEmail, subject, body);

    // Record the alert to prevent duplicates.
    alertLog.appendRow([
      'ALERT-' + Utilities.getUuid(),
      cycleId,
      'Manager Overdue Alert',
      managerEmail,
      now,
      'Sent',
      '',
      '',
      'Awaiting Acknowledgement'
    ]);

    alertedCycles.add(cycleId);
    alertsSent++;
  });

  console.log(
  'Manager alert check complete. Emails sent: ' + alertsSent);
}

function setupManagerAlertTrigger() {
  const functionName = 'sendManagerOverdueAlerts';

  // Remove existing triggers for this function
  // so we don't accidentally schedule duplicates.
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Check for overdue tasks every 5 minutes.
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyMinutes(5)
    .create();

  console.log(
  'Manager alert check complete. Emails sent: ' + alertsSent
);
}


function acknowledgeSelectedAlert() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Alert_Log');

  if (ss.getActiveSheet().getName() !== 'Alert_Log') {
    ui.alert('Open Alert_Log and select the alert row first.');
    return;
  }

  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    ui.alert('Select an alert row below the header.');
    return;
  }

  const data = sheet.getRange(row, 1, 1, 9).getValues()[0];

  if (!data[0]) {
    ui.alert('This row does not contain an alert.');
    return;
  }

  if (data[6]) {
    ui.alert('This alert has already been acknowledged.');
    return;
  }

  const email = Session.getActiveUser().getEmail();

  sheet.getRange(row, 7, 1, 3).setValues([[
    new Date(),
    email || 'Manager',
    'Acknowledged - Pending Resolution'
  ]]);

  ui.alert('Alert acknowledged successfully.');
}

function sendLeadershipEscalations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('Task_Cycles_Log');
  const alertLog = ss.getSheetByName('Alert_Log');

  const leadershipEmail = 'dishanaveen.in@gmail.com';
  const now = new Date();

  const cycles = log.getDataRange().getValues();
  const alerts = alertLog.getDataRange().getValues();

  // Find cycles that have already been escalated.
  const escalatedCycles = new Set();

  alerts.slice(1).forEach(row => {
    if (row[1] && row[2] === 'Leadership Escalation') {
      escalatedCycles.add(row[1]);
    }
  });

  // Load task details once.
  const master = ss.getSheetByName('Task_Master');
  const tasks = master.getDataRange().getValues();

  const taskMap = new Map();

  tasks.slice(1).forEach(task => {
    if (task[0]) taskMap.set(task[0], task);
  });

  let escalationsSent = 0;

  cycles.slice(1).forEach(cycle => {
    const cycleId = cycle[0];
    const taskId = cycle[1];
    const dueDate = cycle[3];
    const status = cycle[4];

    if (!cycleId || !taskId) return;
    if (status === 'Completed') return;
    if (escalatedCycles.has(cycleId)) return;

    if (!(dueDate instanceof Date) || isNaN(dueDate.getTime())) {
      return;
    }

    // Escalate 2 hours and 15 minutes after the deadline.
    const escalationTime = new Date(
      dueDate.getTime() + 135 * 60 * 1000
    );

    if (now < escalationTime) return;

    const task = taskMap.get(taskId);
    if (!task) return;

    const taskName = task[1] || taskId;
    const employee = task[2] || 'Assigned employee';

    const subject = 'Leadership Escalation: ' + taskName;

    const body =
      'Hello,\n\n' +
      'A recurring task remains incomplete beyond its escalation deadline.\n\n' +
      'Task: ' + taskName + '\n' +
      'Task ID: ' + taskId + '\n' +
      'Cycle ID: ' + cycleId + '\n' +
      'Assigned to: ' + employee + '\n' +
      'Original due date: ' + dueDate + '\n\n' +
      'The task requires review and resolution.\n\n' +
      '— Allo Innoware Task Tracker';

    GmailApp.sendEmail(leadershipEmail, subject, body);

    // Log the escalation to prevent duplicate escalation emails.
    alertLog.appendRow([
      'ALERT-' + Utilities.getUuid(),
      cycleId,
      'Leadership Escalation',
      leadershipEmail,
      now,
      'Sent',
      '',
      '',
      'Escalated'
    ]);

    escalatedCycles.add(cycleId);
    escalationsSent++;
  });

  console.log(
    'Leadership escalation check complete. Emails sent: ' +
    escalationsSent
  );
}

function setupLeadershipEscalationTrigger() {
  const functionName = 'sendLeadershipEscalations';

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getUi().alert(
    'Automatic leadership escalation is enabled.'
  );
}

