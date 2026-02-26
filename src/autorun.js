Office.onReady(function() { //waits for office to be ready before running the code, which is important for accessing the Office APIs

var DATA_URL = "https://NDuggan05.github.io/outlook-signature-addin/data/signatures.json";
var LOGO_URL = "https://NDuggan05.github.io/outlook-signature-addin/assets/logo.gif";
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

  // **LOCAL TESTING** var response = await fetch(DATA_URL);  grabs JSON database of employee info from GitHub Pages
  var response = await fetch("https://connacht-signatures.azurewebsites.net/api/signature?email=" + encodeURIComponent(userEmail)); // **PRODUCTION** fetches employee data from an Azure Function API, which looks up the email in the same JSON database but allows for better security and faster lookups than fetching the entire JSON file to the client

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
    // ── Row 1: Name/Title + Contact Details ──
    + '<table cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#333333;line-height:1.5;">'
    + '<tr>'

    // Left: Name & Title
    + '<td style="padding-right:20px;vertical-align:top;">'
    + (emp.name ? '<strong style="font-size:14px;color:#000000;">' + emp.name + '</strong><br/>' : '')
    + (emp.title ? '<span style="font-size:12px;color:#666666;">' + emp.title + '</span>' : '')
    + '</td>'

    // Right: Contact Details
    + '<td style="padding-left:20px;vertical-align:top;border-left:1px solid #cccccc;">'
    + (emp.email ? '<span style="padding-left:10px;"><strong>E:</strong> <a href="mailto:' + emp.email + '" style="color:#333333;text-decoration:underline;">' + emp.email + '</a></span><br/>' : '')
    + (emp.phone ? '<span style="padding-left:10px;"><strong>T:</strong> <a href="tel:' + emp.phone + '" style="color:#333333;text-decoration:underline;">' + emp.phone + '</a></span><br/>' : '')
    + (emp.website ? '<span style="padding-left:10px;"><strong>W:</strong> <a href="https://' + emp.website + '" target="_blank" style="color:#333333;text-decoration:underline;">' + emp.website + '</a></span><br/>' : '')
    + (emp.address ? '<span style="padding-left:10px;"><strong>A:</strong> ' + emp.address + '</span>' : '')
    + '</td>'

    + '</tr>'
    + '</table>'

    // ── Row 2: Banner GIF ──
    + '<table cellpadding="0" cellspacing="0" border="0" style="padding-top:15px;">'
    + '<tr>'
    + '<td>'
    + '<a href="' + WEBSITE_URL + '" target="_blank" style="text-decoration:none;">'
    + (emp.banner ? '<img src="https://NDuggan05.github.io/outlook-signature-addin/assets/' + emp.banner + '" alt="Connacht Hospitality Group" width="500" style="border:0;display:block;" />' : '')
    + '</a>'
    + '</td>'
    + '</tr>'
    + '</table>'

    // ── Row 3: Disclaimer ──
    + '<table cellpadding="0" cellspacing="0" border="0" style="padding-top:15px;">'
    + '<tr>'
    + '<td style="font-size:10px;color:#999999;line-height:1.4;">'
    + '<strong>Disclaimer:</strong><br/><br/>'//Disclaimer title, every plus after this one is a sentence with <br></br> for paragraph spaces
    + 'This email and any attachments may be confidential and intended only for the named recipient. '
    + 'If you receive this email or any attachment(s) in error, please contact the sender by return email and delete it. Thank you.<br/><br/>'
    + 'The sender respects your right to disconnect and does not expect a response outside of your normal working hours unless urgent or pre-agreed.'
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
Office.actions.associate("onNewMessageCompose", onNewMessageCompose); //associate the event handler with the OnNewMessageCompose event so it runs automatically when a new email is composed
}); //end of Office.onReady()
// https://aka.ms/olksideload --> access add ins