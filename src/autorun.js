var DATA_URL = "https://NDuggan05.github.io/outlook-signature-addin/data/signatures.json";
var LOGO_URL = "https://NDuggan05.github.io/outlook-signature-addin/assets/logo.png";
var WEBSITE_URL = "https://www.connachthospitalitygroup.ie/";

var SOCIAL = {
  linkedin:  "https://www.linkedin.com/company/connacht-hospitality-group/",
  facebook:  "https://www.facebook.com/TheConnachtHotel/",
  instagram: "https://www.instagram.com/theconnachthotel/",
};

/* global Office, console */
/*
 * Connacht Hospitality Group — Outlook Signature Add-in
 * autorun.js — Fetches employee data and sets the signature
 */
// ── Logging Helpers ──────────────────────────────────────
var LOG = "[ConnachtSig]";
//each of these functions prepends the LOG tag to the message for easier debugging
function logInfo(msg)  { console.log(LOG, "INFO:", msg); }
function logWarn(msg)  { console.warn(LOG, "WARN:", msg); }
function logError(msg) { console.error(LOG, "ERROR:", msg); }

// ── Notification Helper ──────────────────────────────────
// Displays a notification message in the Outlook UI/ an info bar at the top of the compose window
function notifyUser(type, message) {
  var item = Office.context.mailbox.item; //current email in compose window
  if (!item) return;

  var notificationType =
    type === "error"
      ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
      : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

  item.notificationMessages.replaceAsync("connachtSigStatus", { //replace same notification if calledd multiple times
    type: notificationType,
    message: message,
    persistent: false, // Auto-dismiss after a few seconds
  });
}

// ── Fetch employee data from JSON file ───────────────────
async function getEmployeeData() {
  var userEmail = Office.context.mailbox.userProfile.emailAddress; //Office gives you the current users email
  logInfo("Current user email: " + userEmail);

  var response = await fetch(DATA_URL); // grabs JSON database of employee info from GitHub Pages

  if (!response.ok) {
    throw new Error("Failed to fetch signatures.json — status " + response.status);
  }

  var allEmployees = await response.json(); //parses response as JSON, which should be an object mapping emails to employee data

  // Look up the current user by email (case-insensitive)
  var emailLower = userEmail.toLowerCase();
  var employee = null;

  //loop through all keys in the JSON object to find a case-insensitive match for the user's email
  for (var key in allEmployees) {
    if (key.toLowerCase() === emailLower) {
      employee = allEmployees[key];
      employee.email = key; // Store the email too
      break;
    }
  }

  return employee;
}

// ── Build the signature HTML ─────────────────────────────
function buildSignatureHtml(emp) {
  var html = ''
    + '<table cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333333;line-height:1.4;">'
    + '<tr>'

    // Left column: Logo
    + '<td style="padding-right:15px;vertical-align:top;border-right:2px solid #1a3c34;">'
    + '<a href="' + WEBSITE_URL + '" target="_blank" style="text-decoration:none;">'
    + '<img src="' + LOGO_URL + '" alt="Connacht Hospitality Group" width="150" style="border:0;display:block;" />'
    + '</a>'
    + '</td>'

    // Right column: Details
    + '<td style="padding-left:15px;vertical-align:top;">'

    // Name
    + '<table cellpadding="0" cellspacing="0" border="0">'
    + '<tr><td style="font-size:16px;font-weight:bold;color:#1a3c34;padding-bottom:2px;">'
    + emp.name
    + '</td></tr>'

    // Job title
    + '<tr><td style="font-size:13px;color:#666666;padding-bottom:8px;">'
    + emp.title
    + '</td></tr>'

    // Phone
    + '<tr><td style="font-size:12px;color:#333333;padding-bottom:3px;">'
    + '&#128222; ' + emp.phone
    + '</td></tr>'

    // Email
    + '<tr><td style="font-size:12px;color:#333333;padding-bottom:8px;">'
    + '&#9993; <a href="mailto:' + emp.email + '" style="color:#1a3c34;text-decoration:none;">' + emp.email + '</a>'
    + '</td></tr>'

    // Social links
    + '<tr><td style="padding-top:4px;">'
    + '<a href="' + SOCIAL.linkedin + '" target="_blank" style="text-decoration:none;margin-right:8px;font-size:12px;color:#1a3c34;">LinkedIn</a>'
    + '<a href="' + SOCIAL.facebook + '" target="_blank" style="text-decoration:none;margin-right:8px;font-size:12px;color:#1a3c34;">Facebook</a>'
    + '<a href="' + SOCIAL.instagram + '" target="_blank" style="text-decoration:none;font-size:12px;color:#1a3c34;">Instagram</a>'
    + '</td></tr>'

    + '</table>'
    + '</td>'
    + '</tr>'
    + '</table>';

  return html;
}

// ── Event Handler: OnNewMessageCompose ───────────────────
// This function runs automatically when the user opens a new email compose window
async function onNewMessageCompose(event) {
  logInfo("OnNewMessageCompose triggered");

  // Check that setSignatureAsync is supported
  if (!Office.context.requirements.isSetSupported("Mailbox", "1.10")) { //stops if setSignatureAsync() isnt supported
    logWarn("Mailbox 1.10 not supported");
    notifyUser("informational", "Your Outlook version doesn't support automatic signatures.");
    event.completed();
    return;
  }

  try {
    var employee = await getEmployeeData(); //fetch employee data based on the current user's email

    if (!employee) {
      logWarn("No matching employee found in signatures.json");
      notifyUser("informational", "No signature found for your account. Contact IT to get set up.");
      event.completed();
      return;
    }

    logInfo("Employee found: " + employee.name);
    var signatureHtml = buildSignatureHtml(employee);

    Office.context.mailbox.item.body.setSignatureAsync(
      signatureHtml,
      { coercionType: Office.CoercionType.Html },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          logError("setSignatureAsync failed: " + asyncResult.error.message);
          notifyUser("error", "Could not set signature: " + asyncResult.error.message);
        } else {
          logInfo("Signature set successfully");
        }
        event.completed();
      }
    );

  } catch (error) {
    logError("Error: " + error.message);
    notifyUser("error", "Could not load signature. Check your connection.");
    event.completed();
  }
}

//registers with office
Office.actions.associate("OnNewMessageCompose", onNewMessageCompose); //associate the event handler with the OnNewMessageCompose event so it runs automatically when a new email is composed

