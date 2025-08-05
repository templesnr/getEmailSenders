// Google Apps Script for counting emails by sender in batches. Stopps and resumes after 5mins of processing to avoid reaching script runtime limit.
// 1. Allow required permissions
// 2. Add gmail service to project
// 3. Fill <your-email> in `runProcessor` function and adjust query if needed
// 4. Run function `runProcessor`
// 5. Result will be saved in a new spreadsheet

// Based on ideas from https://stackoverflow.com/a/59222719/5162536 and 
// https://medium.com/geekculture/bypassing-the-maximum-script-runtime-in-google-apps-script-e510aa9ae6da

const execution_threshold = 4 * 60 * 1000 // 4 minutes in ms
const email_batch = 200 // amount of emails to list in one call
const progress_filename = 'emails-by-sender-processor.json'
const progress_fileid_property = 'emails-by-sender-processor-file-id'
const progress_status_property = 'emails-by-sender-processor-status'
const trigger_id_property = 'emails-by-sender-trigger-id'

class SenderList {
  constructor(query, email) {
    if (SenderList.instance) return SenderList.instance;

    this.mailsBySender = new Map();

    this._totalMessages = 0;

    this._query = query;

    this._pageToken = null;

    this._email = email;

    SenderList.instance = this;
    return SenderList.instance;
  }

  get query() {
    return this._query;
  }

  addSender(sender) {
    if (!this.mailsBySender) {
      this.mailsBySender = new Map();
    }

    if (!this.mailsBySender.get(sender)) {
      this.mailsBySender.set(sender, 1)
    } else {
      this.mailsBySender.set(sender, this.mailsBySender.get(sender)+1)
    }

    return this;
  }

  getSenders() {
    return this.mailsBySender;
  }

  addTotalMessages(count) {
    if (!this._totalMessages) {
      this._totalMessages = 0;
    }

    this._totalMessages += count;
    return this;
  }

  set totalMessages(value) {
    this._totalMessages = value;
    return this;
  }

  get totalMessages() {
    return this._totalMessages;
  }

  set pageToken(token) {
    this._pageToken = token;
    return this;
  }

  get pageToken() {
    return this._pageToken;
  }

  get email() {
    return this._email;
  }

  toJSON() {
    return {
      mailsBySender: JSON.stringify([...this.mailsBySender.entries()]),
      totalMessages: this.totalMessages,
      query: this.query,
      pageToken: this.pageToken,
      email: this.email,
    };
  }

  import(json) {
    this.mailsBySender = new Map(JSON.parse(json.mailsBySender));
    this._totalMessages = json.totalMessages;
    this._query = json.query;
    this._pageToken = json.pageToken;
    this._email = json.email;
    return this;
  }
}

class Timer {
  start() {
    this.start = Date.now();
  }

  getDuration() {
    return Date.now() - this.start;
  }
}

class ProcessorStatus {
  static set(status) {
    PropertiesService.getScriptProperties().setProperty(progress_status_property, status);
  }

  static get() {
    return PropertiesService.getScriptProperties().getProperty(progress_status_property);
  }
}

class Trigger {
  constructor(functionName, minutesTick) {
    let trigger = ScriptApp.newTrigger(functionName).timeBased().everyMinutes(minutesTick).create();
    PropertiesService.getScriptProperties().setProperty(trigger_id_property, trigger.getUniqueId());
    return trigger;
  }

  static isAlreadyCreated() {
    let triggerId = PropertiesService.getScriptProperties().getProperty(trigger_id_property);

    if (triggerId) {
      let triggers = ScriptApp.getProjectTriggers();
      let existingTrigger = triggers.find((trigger) => trigger.getUniqueId() === triggerId && trigger.getHandlerFunction() === 'runProcessor');
      if (existingTrigger) return true;
    }

    return false;
  }

  static deleteTrigger(e) {
    if (typeof e !== 'object') return Logger.log(`${e} is not an event object`);
    if (!e.triggerUid) return Logger.log(`${JSON.stringify(e)} doesnt have a triggerUid`);

    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getUniqueId() === e.triggerUid) return ScriptApp.deleteTrigger(trigger);
    });
  }
}

function runProcessor(e) {
  try {
    var timer = new Timer();
    timer.start();

    if (ProcessorStatus.get() === 'running') return Logger.log("exiting because processor already running");

    ProcessorStatus.set('running')

    let query = "in: inbox"; // query for filtering emails
    let email = 'templesnr@gmail.com' // your email here
    var senderList = new SenderList(query, email);

    const existingProgressFileId = PropertiesService.getScriptProperties().getProperty(progress_fileid_property);
    if (existingProgressFileId) {
      Logger.log('resuming processing...');
      let json = readProgressFile(existingProgressFileId);
      if (json) {
        senderList = senderList.import(json);
      } else {
        Logger.log(`progress file with id ${existingProgressFileId} does not exist, starting processing from beginning...`)
        PropertiesService.getScriptProperties().deleteProperty(progress_fileid_property);
      }
    }

    Logger.log(`starting processing sender list for query '${senderList.query}' and email '${senderList.email}'. Current total messages processed: ${senderList.totalMessages}`)

    return processSenderList(senderList, timer, e);
  } catch(e) {
    Logger.log(`error occured while processing, setting processor status to failed: ${e}`);
    return ProcessorStatus.set('failed');
  }
}

function processSenderList(senderList, timer, e) {
  do {
    if (timer.getDuration() >= execution_threshold) { 
      Logger.log('stopping because execution threshold reached, saving progress and creating trigger...');
      writeProgressFile(senderList);
      ProcessorStatus.set('not running');
      if (!Trigger.isAlreadyCreated()) new Trigger('runProcessor', 1);
      return;
    }

    var result = Gmail.Users.Messages.list(senderList.email, { maxResults: email_batch, pageToken: senderList.pageToken, q: senderList.query });

    if (result.messages) {
      for (var i = 0; i < result.messages.length; i++) {
        var sender = GmailApp.getMessageById(result.messages[i].id).getFrom();
        senderList.addSender(sender);
        senderList.addTotalMessages(1);
      }

      Logger.log('current total messages: ' + senderList.totalMessages)
    }
     
    senderList.pageToken = result.nextPageToken
  } while (senderList.pageToken);

  Logger.log('finished, got total messages: ' + senderList.totalMessages)
  writeSenderList(senderList);
  cleanUp(e);
}

function cleanUp(e) {
  Logger.log('cleaning up...')
  Trigger.deleteTrigger(e);
  
  const existingProgressFileId = PropertiesService.getScriptProperties().getProperty(progress_fileid_property);
  if (existingProgressFileId) {
    let file = DriveApp.getFileById(existingProgressFileId);
    if (file.getName() === progress_filename) {
      file.setTrashed(true);
    } else {
      Logger.log(`not deleting progress file because unexpected name of file: '${file.getName()}'`)
    }
  }

  PropertiesService.getScriptProperties().deleteProperty(progress_fileid_property);
  PropertiesService.getScriptProperties().deleteProperty(progress_status_property);
  PropertiesService.getScriptProperties().deleteProperty(trigger_id_property);
}

function writeProgressFile(senderList) {
  let file;
  const existingFileId = PropertiesService.getScriptProperties().getProperty(progress_fileid_property);
  if (existingFileId) {
    file = DriveApp.getFileById(existingFileId);
  } else {
    file = DriveApp.createFile(progress_filename, '');
    PropertiesService.getScriptProperties().setProperty(progress_fileid_property, file.getId());
  }

  file.setContent(JSON.stringify(senderList));
}

function readProgressFile(fileId) {
  try {
    let file = DriveApp.getFileById(fileId);
    if (file.isTrashed()) return null;
    return JSON.parse(file.getBlob().getDataAsString());
  } catch (e) {
    Logger.log(`warning: failed to open progress file with id ${fileId}: ${e}`);
    return null;
  }
}

function writeSenderList(senderList) {
  var sender_array = Array.from(senderList.getSenders(), ([sender, count]) => ([ sender, count ]));

  var ss = SpreadsheetApp.create(`Gmail count emails by sender for query '${senderList.query}' (${new Date()})`);
  var sh = ss.getActiveSheet()
  sh.clear();

  Logger.log(`writing to spreadsheet '${ss.getSheetName()}' of document '${ss.getName()}'`)
  sh.appendRow(['Email Address', 'Count']);
  sh.getRange(2, 1, sender_array.length, 2).setValues(sender_array).sort({ column: 2, ascending: false });
}
