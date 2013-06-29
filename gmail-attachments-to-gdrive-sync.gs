/*
 *
 * The MIT License (MIT)
 *
 * Copyright (c) 2013 Avi Dullu
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */

var kGmailAttachmentFolder = "GmailAttachments";

function MyLogger() {
  this.log = function(data) {
    if (kEnableDebugging) {
      Logger.log(data);
    }
  }
  this.errorLog = function(data) {
    Logger.log(data);
  }
}
var kGlobalLogger = new MyLogger();

function getAllowedContentTypes() {
  var allowedContentTypes = new BucketsLib.buckets.Set();
  // Text
  allowedContentTypes.add("text/plain");
  allowedContentTypes.add("text/csv");
  allowedContentTypes.add("text/xml");
  allowedContentTypes.add("text/html");
  allowedContentTypes.add("text/cmd");
  allowedContentTypes.add("text/javascript");
  // Images
  allowedContentTypes.add("image/jpeg");
  allowedContentTypes.add("image/gif");
  allowedContentTypes.add("image/tiff");
  allowedContentTypes.add("image/svg+xml");
  // Application
  allowedContentTypes.add("application/pdf");
  allowedContentTypes.add("application/javascript");
  allowedContentTypes.add("application/json");
  allowedContentTypes.add("application/xml");
  allowedContentTypes.add("application/zip");
  allowedContentTypes.add("application/msword");
  allowedContentTypes.add("application/vnd.ms-powerpoint");
  allowedContentTypes.add("application/x-shockwave-flash");
  allowedContentTypes.add("application/vnd.ms-project");
  allowedContentTypes.add("application/vnd.ms-works");
  allowedContentTypes.add("application/vnd.ms-outlook");
  allowedContentTypes.add("application/vnd.ms-excel");
  allowedContentTypes.add("application/postscript");
  allowedContentTypes.add("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  allowedContentTypes.add("application/vnd.openxmlformats-officedocument.presentationml.presentation");
  allowedContentTypes.add("application/vnd.mozilla.xul+xml");
  allowedContentTypes.add("application/vnd.oasis.opendocument.spreadsheet");
  allowedContentTypes.add("application/vnd.oasis.opendocument.presentation");
  allowedContentTypes.add("application/vnd.oasis.opendocument.text");
  
  // Audio
  allowedContentTypes.add("audio/mpeg");
  allowedContentTypes.add("audio/mp4");
  allowedContentTypes.add("audio/mid");
  
  // Video
  allowedContentTypes.add("video/mpeg");
  allowedContentTypes.add("video/mp4");
  allowedContentTypes.add("video/x-flv");
  allowedContentTypes.add("video/x-ms-wmv");
  allowedContentTypes.add("video/x-msvideo");
  allowedContentTypes.add("video/quicktime");

  kGlobalLogger.log("Total allowed content types are "
                    + allowedContentTypes.size());

  return allowedContentTypes;
}
var kAllowedContentTypes = getAllowedContentTypes();

var kEnableDebugging = false;
////////////////////////////////////////////////////////////////////////////
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

var kFolder = "f";
var kLastSynced = "l";
var kFilesAddedToday = "fa";
var kUsersCount = "ku";
var kMaxFilesAddPerDay = 240;

var kOneMinute = 60000;
var kOneDay = 86400000;
var kThirtyDays = 2592000000;
var kOneYear = 31622400000;
var k6Hours = 21600000;

function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; ++i) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function processInput(form) {
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    var count = parseInt(ScriptProperties.getProperty(kUsersCount));
    ++count;
    ScriptProperties.setProperty(kUsersCount, count.toString());
    lock.releaseLock();
  } catch (e) {
    // Its ok, could not add this user.
  }
  // The user is trying to re-run the script. Flush out all past data,
  // the files in 'GmailAttachments' will not be affected.
  UserProperties.deleteAllProperties();
  deleteAllTriggers();

  UserProperties.setProperty(kFolder, form.folderStructure);
  var syncYears = parseInt(form.syncFrom);
  var today = Date.parse(new Date().toJSON());
  today = today - syncYears * kOneYear;
  UserProperties.setProperty(kLastSynced, new Date(today).toJSON());
  UserProperties.setProperty(kFilesAddedToday,
      Utilities.jsonStringify({ 'added' : 0,
                                'date' : new Date().toJSON() }));

  // One time trigger, so that we return to user early
  // and start syncing in background
  ScriptApp.newTrigger("syncUserDrives").timeBased().after(2 * kOneMinute).create();
  var reply = "<br/><br/><br/>The script was successful and will start syncing your attachments to your drive within a minute!";
  reply += "<br/> <br/>Due to Google App Script policies, we can't add more than 250 files each day to GDrive, ";
  reply += "<br/> and a limit on the CPU processing which can be done in one day, ";
  reply += "<br/>so the initial sync will take a couple of days, but we'll eventually get there!";
  reply += "<br/> Google App Script system will have sent you an email describing how to disable this script, ";
  reply += "<br/>if you do decide to do so, please let me know what caused you trouble and I'll try to improve this. ";
  reply += "<br/>If you get a new attachment, it should show up in your GDrive within 5-10 minutes. ";
  reply += "<br/>You should also consider installing the GDrive for your PC/Mac and link it to this account. ";
  reply += "<br/>By doing that, all these attachments will automatically sync to your laptop and you need not download ";
  reply += "<br/> them yourself.<br/><br/> Thanks again for installing! <br/><br/>Avi";
  return reply;
}

function UserData(structure, lastSynced) {
  this.structure = structure;
  this.lastSynced = lastSynced;
}

function HasTime(start) {
  return Date.parse(new Date().toJSON()) - Date.parse(start.toJSON()) < kOneMinute;
}

function syncUserDrives(e) {
  var today = new Date();
  // First, delete the trigger which triggered this. There is always
  // just one trigger active for each user.
  deleteAllTriggers();
  // We make multiple trigers so that if App Script time trigger screws up, we are not screwed
  ScriptApp.newTrigger("syncUserDrives").timeBased().after(k6Hours).create();
  ScriptApp.newTrigger("syncUserDrives").timeBased().after(2 * k6Hours).create();
  ScriptApp.newTrigger("syncUserDrives").timeBased().after(4 * k6Hours).create();
  ScriptApp.newTrigger("syncUserDrives").timeBased().after(8 * k6Hours).create();
  ScriptApp.newTrigger("syncUserDrives").timeBased().after(28 * k6Hours).create();

  var filesAddedToday = Utilities.jsonParse(
      UserProperties.getProperty(kFilesAddedToday));
  var added = parseInt(filesAddedToday['added']);
  var last = new Date(Date.parse(filesAddedToday['date'])).getDate();
  if (last == today.getDate() && added >= kMaxFilesAddPerDay) {
    kGlobalLogger.log("Used up all quota for today. Can't add more files.");
    ScriptApp.newTrigger("syncUserDrives").timeBased().after(60 * kOneMinute).create();
    return;
  }
  var user = new UserData(UserProperties.getProperty(kFolder),
                          UserProperties.getProperty(kLastSynced));
  if (last != today.getDate()) {
    added = 0;
  }
  var fullSynced = false;
  var someError = false;
  try {
    while (HasTime(today) &&
           added < kMaxFilesAddPerDay &&
           !fullSynced) {
      var maxToAdd = kMaxFilesAddPerDay - added;
      var searchQuery = getSearchQuery(user);
      kGlobalLogger.log("Query : " + searchQuery['query']
                        + "  files asking: " + maxToAdd);
      var threads = GmailApp.search(searchQuery['query']);
      if (threads.length > 0) {
        var filesAdded = updateDrive(threads, user, maxToAdd);
        if (filesAdded['has_error'] == "t") {
          // We got some error, should backoff and come back after some time
          someError = true;
        }
        added += filesAdded['addedFiles'].size();
      }
      fullSynced = (searchQuery['isToday'] == "true");
      if (added < kMaxFilesAddPerDay) {
        if (fullSynced) {
          user.lastSynced = today.toJSON();
        } else {
          user.lastSynced = new Date(Date.parse(user.lastSynced) + kThirtyDays).toJSON();
        }
      } else {
        // Let us keep the last synced as is.
      }
    }
  } catch (e) {
    kGlobalLogger.errorLog("Some error in execution: " + e );
  }
  UserProperties.setProperty(kLastSynced, user.lastSynced);
  UserProperties.setProperty(kFilesAddedToday, Utilities.jsonStringify(
      {'added' : added, 'date' : today.toJSON()}));
  if (!someError && fullSynced) {
    ScriptApp.newTrigger("syncUserDrives").timeBased().after(12 * kOneMinute).create();
  } else {
    ScriptApp.newTrigger("syncUserDrives").timeBased().after(20 * kOneMinute).create();
  }
}

function getSearchQuery(userData) {
  var query = "to:me in:all has:attachment -in:trash -in:spam ";
  var nextDate = kThirtyDays + Date.parse(userData.lastSynced);
  var today = Date.parse(new Date().toJSON());
  var isToday = "false";
  if (today < nextDate) {
    isToday = "true";
    nextDate = today;
  }
  nextDate += kOneDay;
  var afterDate = new Date(Date.parse(userData.lastSynced) - kOneDay);
  var beforeDate = new Date(nextDate);
  // JS returns 0..11, Gmail uses 1..12
  var b_tmp = beforeDate.getMonth() + 1;
  var a_tmp = afterDate.getMonth() + 1;
  var beforeQ = "" + beforeDate.getFullYear() + "/" + b_tmp + "/" + beforeDate.getDate() + " ";
  var afterQ  = "" + afterDate.getFullYear() +  "/" + a_tmp + "/" + afterDate.getDate()  + " ";
  if (beforeQ == afterQ) {
    var yester = new Date(Date.parse(userData.lastSynced) - kOneDay);
    a_tmp = yester.getMonth() + 1;
    afterQ  = "" + yester.getFullYear() +  "/" + a_tmp + "/" + yester.getDate()  + " ";
    query = query + "  " + "after:" + afterQ;
  } else {
    query = query + "before:" + beforeQ + "  " + "after:" + afterQ;
  }
  return { 'query' : query, 'before' : beforeDate.toJSON(), "isToday" : isToday };
}

function getOrCreateFolderInFolder(parentFolder, newFolderName) {
  var folders = parentFolder.getFolders();
  var cf = null;
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() == newFolderName) {
      cf = folder;
      kGlobalLogger.log("Found existing folder " + folder.getName());
      break;
    }
  }
  if (cf == null) {
    kGlobalLogger.log("Creating new folder " + newFolderName);
    cf = parentFolder.createFolder(newFolderName);
  }
  return cf;
}

function getAttachmentName(fileName, from, date) {
  var dateString = date.toDateString().split(" ");
  return from + "-"
      + dateString[2] + dateString[1] + dateString[3] + "-"
      + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds() + "-"
      + fileName;
}

// Give the date of the attachment, returns a folder name like
// Jun2012
function getFolderNameForMonthYear(date) {
  var s = date.toDateString().split(" ");
  var ret_val = s[1] + s[3];
  kGlobalLogger.log("Folder name for date " + date.toDateString()
                    + " is " + ret_val);
  return ret_val;
}

function getFolderNameForContentType(type) {
  var content = type.split("/")[0];
  var ret = "";
  switch (type.split("/")[0]) {
    case "text":
      ret = "Text Data"; break;
    case "image":
      ret = "Images"; break;
    case "application":
      ret = "Application Data"; break;
    case "audio":
      ret = "Audio"; break;
    case "video":
      ret = "Video"; break;
    default:
      // We will not store this, empty return value implies this.
      kGlobalLogger.errorLog("Invalid content: " + type);
  }
  kGlobalLogger.log("Folder name for " + type + " is " + ret);
  return ret;
}

// We do not need the actual file here because the structure
// user chose is fixed
function getFolderNameForFile(contentType, date, userData, existingFolders) {
  var folderName = "";
  if (userData.structure == "byContentType") {
    folderName = getFolderNameForContentType(contentType);
  } else if (userData.structure == "byMonthly") {
    folderName = getFolderNameForMonthYear(date);
  } else if (userData.structure == "byYear") {
    folderName = "" + date.getFullYear();
  } else {
    kGlobalLogger.errorLog("Invalid structure type " + userData.structure);
  }
  if (folderName != "") {
    if (existingFolders.containsKey(folderName)) {
        return existingFolders.get(folderName);
    } else {
      var headFolder = getOrCreateFolderInFolder(DriveApp, kGmailAttachmentFolder);
      var newFolder = headFolder.createFolder(folderName);
      existingFolders.set(folderName, newFolder);
      return newFolder;
    }
  }
  return null;
}

// TODO: Add a last run error log to UserProperties
function updateDrive(gmailThreads, userData, maxToAdd) {
  // Check if we need the top level folder.
  var folder = getOrCreateFolderInFolder(DriveApp, kGmailAttachmentFolder);

  // Build a map of existing folders, we will create new on the fly if we need to.
  // map<FolderName, Folder*>
  var existingFolders = new BucketsLib.buckets.Dictionary();
  var currFolders = folder.getFolders();
  while (currFolders.hasNext()) {
    var currFolder = currFolders.next();
    existingFolders.set(currFolder.getName(), currFolder);
    kGlobalLogger.log("Found folder " + currFolder.getName());
  }
  kGlobalLogger.log("Already present folders :" + existingFolders.size());

  // map<NewAddedFileName, Folder*>
  var addedFiles = new BucketsLib.buckets.Dictionary();
  var addMore = (maxToAdd > 0);
  var has_error = "f";
  for (var i = 0; (i < gmailThreads.length) && addMore; ++i) {
    var thread = gmailThreads[i];
    var messages = thread.getMessages();
    for (var j = 0; j < messages.length && addMore; ++j) {
      var message = messages[j];
      var attachments = message.getAttachments();
      for (var k = 0; k < attachments.length && addMore; ++k) {
        var attachment = attachments[k];
        var contentType = attachment.getContentType();
        var fileName = getAttachmentName(attachment.getName(),
                                         message.getFrom(),
                                         message.getDate());
        if (kAllowedContentTypes.contains(contentType)
            && !addedFiles.containsKey(fileName)) {
          kGlobalLogger.log("Accepted file of type: " + contentType
                            + "  name: " + fileName);
          var folderForFile = getFolderNameForFile(contentType,
                                                   message.getDate(),
                                                   userData,
                                                   existingFolders);
          if (folderForFile == null) {
            kGlobalLogger.errorLog("Could not find directory.");
            // If some error, we set all the attachments in the top level folder
            // Can be used to recover later.
            folderForFile = getOrCreateFolderInFolder(DriveApp,
                                                      kGmailAttachmentFolder);
          }
          try {
            var iter = folderForFile.getFilesByName(fileName);
            // duplicate.
            if (iter != null && iter.hasNext()) {
              continue;
            }
            var blob = attachment.copyBlob();
            blob.setName(fileName);
            var f = folderForFile.createFile(blob);
            f.setDescription("Created from a Gmail Attachment from " + message.getFrom()
                + " sent on date " + message.getDate());
            addedFiles.set(fileName, folderForFile);
            Utilities.sleep(100);
            if (addedFiles.size() == maxToAdd) addMore = false;
          } catch (e) {
            kGlobalLogger.errorLog("Some problem with an attachment : " + e);
            has_error = "t";
            Utilities.sleep(1000);
          }
        } else {
          kGlobalLogger.errorLog("Rejected file of type: " + contentType);
        }
      }
    }
  }
  return { 'addedFiles' : addedFiles, 'has_error' : has_error };
}

