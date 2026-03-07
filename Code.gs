function sendReminders() {
  const props = PropertiesService.getUserProperties();

  const sheetName = props.getProperty("sheetName");
  const taskCol = Number(props.getProperty("taskCol")) - 1;
  const dueDateCol = Number(props.getProperty("dueDateCol")) - 1;
  const emailCol = Number(props.getProperty("emailCol")) - 1;
  const template =
    props.getProperty("template") ||
    "Reminder: '{{task}}' is due on {{dueDate}}.";

  const reminderType = props.getProperty("reminderType");
  const daysBefore = Number(props.getProperty("daysBefore") || 0);

  if (!sheetName) return;

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const today = normalizeDate(new Date());
  const logSheet = getLogSheet();

  for (let i = 1; i < data.length; i++) {
    const task = data[i][taskCol];
    const dueDateRaw = data[i][dueDateCol];
    const email = data[i][emailCol];

    if (!email || !dueDateRaw) continue;

    const dueDate = normalizeDate(new Date(dueDateRaw));
    const diffDays =
      (dueDate.getTime() - today.getTime()) / (1000 * 3600 * 24);

    let shouldSend = false;

    if (reminderType === "same" && diffDays === 0) shouldSend = true;
    if (reminderType === "before" && diffDays === daysBefore) shouldSend = true;

    if (!shouldSend) continue;

    const formattedDate = Utilities.formatDate(
      dueDate,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );

    const messageBody = template
      .replace("{{task}}", task)
      .replace("{{dueDate}}", formattedDate);

    try {
      MailApp.sendEmail({
        to: email,
        subject: `Reminder: ${task}`,
        body: messageBody,
      });

      logSheet.appendRow([
        new Date(),
        task,
        email,
        formattedDate,
        "Sent",
      ]);
    } catch (err) {
      logSheet.appendRow([
        new Date(),
        task,
        email,
        formattedDate,
        "Failed",
      ]);
    }
  }
}

function saveUserSettings(settings) {
  const props = PropertiesService.getUserProperties();

  props.setProperty("sheetName", settings.sheetName);
  props.setProperty("taskCol", settings.taskCol);
  props.setProperty("dueDateCol", settings.dueDateCol);
  props.setProperty("emailCol", settings.emailCol);
  props.setProperty("template", settings.template);
  props.setProperty("reminderType", settings.reminderType);
  props.setProperty("daysBefore", settings.daysBefore || "0");
  props.setProperty("timeSlots", JSON.stringify(settings.timeSlots));
}

function createTriggersFromSettings() {
  const props = PropertiesService.getUserProperties();
  const slots = JSON.parse(props.getProperty("timeSlots") || "[]");

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(t => {
    if (t.getHandlerFunction() === "sendReminders") {
      ScriptApp.deleteTrigger(t);
    }
  });

  slots.forEach(hour => {
    ScriptApp.newTrigger("sendReminders")
      .timeBased()
      .everyDays(1)
      .atHour(Number(hour))
      .create();
  });
}

function getUserSettings() {
  const props = PropertiesService.getUserProperties();

  return {
    sheetName: props.getProperty("sheetName") || "",
    taskCol: props.getProperty("taskCol") || "1",
    dueDateCol: props.getProperty("dueDateCol") || "2",
    emailCol: props.getProperty("emailCol") || "3",
    template: props.getProperty("template") || "",
    reminderType: props.getProperty("reminderType") || "same",
    daysBefore: props.getProperty("daysBefore") || "1",
    timeSlots: props.getProperty("timeSlots") || "[]"
  };
}

function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName());
}

function getColumnHeaders(sheetName) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function normalizeDate(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Acreminder_Log");

  if (!logSheet) {
    logSheet = ss.insertSheet("Acreminder_Log");
    logSheet.appendRow([
      "Timestamp",
      "Task",
      "Email",
      "Due Date",
      "Status",
    ]);
  }

  return logSheet;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Acreminder")
    .addItem("Open Settings", "showSidebar")
    .addItem("Send Reminders Now", "sendReminders")
    .addToUi();
}

function showSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile("UI").setTitle("Acreminder")
  );
}
