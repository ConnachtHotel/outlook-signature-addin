Office.onReady(function () {

console.log("[ConnachtSig] JS Loaded", new Date().toISOString());
//actual Function App URL from the Azure Portal.
var API_URL = "https://connachtsignatures-bsbfakbbcjf6fnbb.westeurope-01.azurewebsites.net/api/signature";

// Added HOTEL_CONFIG ─────────────────────────
// WHY: Each hotel/brand has different details — banner, website, address, and visual style.
//      When an employee's email is @theconnacht.ie they get the Connacht config.
//      When it's @hydehotel.ie they get the Hyde config. And so on.
//      The Azure Function handles the banner selection too, but this config controls
//      the website link on the banner, the address fallback, and all the styling.
//      To add a new hotel, just add a new entry here.

var HOTEL_CONFIG = {
    "@chgl.ie": {
        website: "www.chgl.ie",
        address: "Connacht Hotel, Old Dublin Rd, Galway, H91 K5DD",
        websiteUrl: "https://www.connachthospitalitygroup.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#000000",
            textColor: "#000000",
            dividerColor: "#000000",
            linkColor: "#000000",
            disclaimerColor: "#000000",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "@theconnacht.ie": {
        website: "www.theconnacht.ie",
        address: "Connacht Hotel, Old Dublin Rd, Galway, H91 K5DD",
        websiteUrl: "https://www.theconnacht.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "@hydehotel.ie": {
        website: "www.hydehotel.ie",
        address: "Forster St, Galway, H91 R2K3",
        websiteUrl: "https://www.hydehotel.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "@theresidencehotel.ie": {
        website: "www.theresidencehotel.ie",
        address: "14 Quay Street, Galway, H91 X580",
        websiteUrl: "https://www.theresidencehotel.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "@anpucan.ie": {
        website: "www.anpucan.ie",
        address: "Forster St, Galway",
        websiteUrl: "https://www.anpucan.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "@activefitness.ie": {
        website: "www.activefitness.ie",
        address: "Old Dublin Rd, Galway",
        websiteUrl: "https://www.activefitness.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "@galwayhooker.ie": {
        website: "www.galwayhooker.ie",
        address: "Galway City",
        websiteUrl: "https://www.galwayhooker.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "@thehawthornhotel.ie": {
        website: "www.thehawthornhotel.ie",
        address: "Hawthorn Hotel, Galway",
        websiteUrl: "https://www.thehawthornhotel.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "@mfitzgeraldsbar.ie": {
        website: "www.mfitzgeraldsbar.ie",
        address: "Galway City",
        websiteUrl: "https://www.mfitzgeraldsbar.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "@connachthospitalitygroup.ie": {
        website: "www.connachthospitalitygroup.ie",
        address: "Connacht Hotel, Old Dublin Rd, Galway, H91 K5DD",
        websiteUrl: "https://www.connachthospitalitygroup.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    },
    "default": {
        website: "www.chgl.ie",
        address: "Connacht Hotel, Old Dublin Rd, Galway, H91 K5DD",
        websiteUrl: "https://www.connachthospitalitygroup.ie/",
        style: {
            nameColor: "#000000",
            nameSize: "14px",
            titleColor: "#666666",
            textColor: "#333333",
            dividerColor: "#cccccc",
            linkColor: "#333333",
            disclaimerColor: "#999999",
            fontFamily: "Arial,Helvetica,sans-serif"
        }
    }
};

/* global Office, console */
/*
 * Connacht Hospitality Group — Outlook Signature Add-in
 * autorun.js — Fetches employee data and sets the signature
 */

// ── Logging Helpers (NO CHANGES) ─────────────────────────
var LOG = "[ConnachtSig]";
function logInfo(msg)  { console.log(LOG, "INFO:", msg); }
function logWarn(msg)  { console.warn(LOG, "WARN:", msg); }
function logError(msg) { console.error(LOG, "ERROR:", msg); }

// ── Notification Helper (NO CHANGES) ─────────────────────
function notifyUser(type, message) {
    var item = Office.context.mailbox.item;
    if (!item) return;

    var notificationType =
        type === "error"
            ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
            : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

    item.notificationMessages.replaceAsync("connachtSigStatus", {
        type: notificationType,
        message: message
    });
}

// ── CHANGE 5: Added getConfigForEmail helper ─────────────
// WHY: Takes an email like "NDuggan@chgl.ie", extracts the suffix "@chgl.ie",
//      and returns the matching HOTEL_CONFIG entry. If no match, returns default.
//      This is how each hotel gets its own style, website, and address.
function getConfigForEmail(email) {
    var emailLower = email.toLowerCase();
    for (var suffix in HOTEL_CONFIG) {
        if (suffix !== "default" && emailLower.endsWith(suffix)) {
            return HOTEL_CONFIG[suffix];
        }
    }
    return HOTEL_CONFIG["default"];
}

// ── CHANGE 6: Completely rewritten getEmployeeData ───────
// OLD: Fetched the entire JSON file, then looped through every employee to find a match.
// NEW: Calls the Azure Function with just the email. The function calls Microsoft Graph,
//      finds the employee in Azure AD, and returns only that one employee's data.
//      No more looping. No more maintaining a JSON file.
//      The response is already in the right format — name, title, phone, email, banner, etc.
async function getEmployeeData() {
    var userEmail = Office.context.mailbox.userProfile.emailAddress;
    logInfo("Current user email: " + userEmail);

    var response = await fetch(API_URL + "?email=" + encodeURIComponent(userEmail));

    // 404 means the employee wasn't found in Azure AD
    if (response.status === 404) {
        logWarn("No employee found in Azure AD for: " + userEmail);
        return null;
    }

    if (!response.ok) {
        throw new Error("API request failed — status " + response.status);
    }

    // The Azure Function returns a single employee object directly
    // No need to loop or search — it's already the right person
    var employee = await response.json();
    return employee;
}

// buildSignatureHtml now takes a config parameter
// OLD: Used hardcoded colours and a hardcoded WEBSITE_URL for the banner link.
// NEW: Takes a config object from HOTEL_CONFIG so every colour, font, and link
//      can be different per hotel. The banner link now goes to the specific hotel's
//      website (config.websiteUrl) instead of a hardcoded URL.
//      Also, emp.banner now contains the full URL from the Azure Function.
function buildSignatureHtml(emp, config) {
    var s = config.style;

    var html = ''
        // ── Row 1: Name/Title + Contact Details ──
        + '<table cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;font-family:' + s.fontFamily + ';font-size:12px;color:' + s.textColor + ';line-height:1.5;">'
        + '<tr>'

        // Left: Name & Title
        + '<td style="padding-right:20px;vertical-align:top;padding-bottom:15px;">'
        + (emp.name ? '<strong style="font-size:' + s.nameSize + ';color:' + s.nameColor + ';">' + emp.name + '</strong><br/>' : '')
        + (emp.title ? '<span style="font-size:12px;color:' + s.titleColor + ';">' + emp.title + '</span>' : '')
        + '</td>'

        // Right: Contact Details
        + '<td style="padding-left:20px;vertical-align:top;border-left:1px solid ' + s.dividerColor + ';">'
        + (emp.email ? '<span style="padding-left:10px;"><strong>E:</strong> <a href="mailto:' + emp.email + '" style="color:' + s.linkColor + ';text-decoration:underline;">' + emp.email + '</a></span><br/>' : '')
        + (emp.phone ? '<span style="padding-left:10px;"><strong>T:</strong> <a href="tel:' + emp.phone + '" style="color:' + s.linkColor + ';text-decoration:underline;">' + emp.phone + '</a></span><br/>' : '')
        + (emp.website ? '<span style="padding-left:10px;"><strong>W:</strong> <a href="https://' + emp.website + '" target="_blank" style="color:' + s.linkColor + ';text-decoration:underline;">' + emp.website + '</a></span><br/>' : '')
        + (emp.address ? '<span style="padding-left:10px;"><strong>A:</strong> ' + emp.address + '</span>' : '')
        + '</td>'

        + '</tr>'

        + '<tr>'
        + '<td colspan="2" style="height:15px;"></td>'
        + '</tr>'

        + '</table>'

        // ── Row 2: Banner GIF ──
        + '<table cellpadding="0" cellspacing="0" border="0" style="padding-top:15px;">'
        + '<tr>'
        + '<td>'
        + '<a href="' + config.websiteUrl + '" target="_blank" style="text-decoration:none;">'
        + (emp.banner ? '<img src="' + emp.banner + '" alt="Connacht Hospitality Group" width="500" style="border:0;display:block;max-width:100%;height:auto;" />' : '')
        + '</a>'
        + '</td>'
        + '</tr>'

        + '<tr>'
        + '<td colspan="2" style="height:15px;"></td>'
        + '</tr>'

        + '</table>'

        // ── Row 3: Disclaimer ──
        + '<table cellpadding="0" cellspacing="0" border="0" style="padding-top:15px;">'
        + '<tr>'
        + '<td style="font-size:10px;color:' + s.disclaimerColor + ';line-height:1.4;">'
        + '<strong>Disclaimer:</strong><br/><br/>'
        + 'This email and any attachments may be confidential and intended only for the named recipient. '
        + 'If you receive this email or any attachment(s) in error, please contact the sender by return email and delete it. Thank you.<br/><br/>'
        + 'The sender respects your right to disconnect and does not expect a response outside of your normal working hours unless urgent or pre-agreed.'
        + '</td>'
        + '</tr>'
        + '</table>';

    return html;
}

async function onNewMessageCompose(event) {
    logInfo("OnNewMessageCompose triggered");

    Office.context.mailbox.item.notificationMessages.addAsync(
  "connachtDebug",
  {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Connacht Signature runtime triggered",
    icon: "icon16",
    persistent: false
  }
);

    // Mobile kills the add-in after 60s, this fires at 55s to exit cleanly
    var safetyTimeout = setTimeout(function() {
        logWarn("Safety timeout reached — completing event early");
        event.completed();
    }, 55000);

    var bodyItem = Office.context.mailbox.item.body;
    if (!bodyItem.setSignatureAsync && !bodyItem.prependAsync) {
        logWarn("Neither setSignatureAsync nor prependAsync supported");
        notifyUser("informational", "Your Outlook version doesn't support automatic signatures.");
        clearTimeout(safetyTimeout);
        event.completed();
        return;
    }

    try {
        var employee = await getEmployeeData();

        if (!employee) {
            logWarn("No matching employee found");
            notifyUser("informational", "No signature found for your account. Contact IT to get set up.");
            clearTimeout(safetyTimeout);
            event.completed();
            return;
        }

        logInfo("Employee found: " + employee.name);

        // NEW: Get the hotel config based on email suffix
        var config = getConfigForEmail(employee.email);

        // NEW: Apply config values as fallbacks
        // Banner always comes from the Azure Function (based on email suffix in signature.js)
        // Website and address use the hotel config if not set on the employee
        employee.website = employee.website || config.website;
        employee.address = employee.address || config.address;

        logInfo("Config applied — website: " + employee.website);

        // CHANGE: Now passes config as second argument for styling
        var signatureHtml = buildSignatureHtml(employee, config);

        if (bodyItem.setSignatureAsync) {
            // Modern clients — sets signature in the correct position automatically
            bodyItem.setSignatureAsync(
                signatureHtml,
                { coercionType: Office.CoercionType.Html },
                function (asyncResult) {
                    clearTimeout(safetyTimeout);
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        logError("setSignatureAsync failed: " + asyncResult.error.message);
                        notifyUser("error", "Could not set signature: " + asyncResult.error.message);
                    } else {
                        logInfo("Signature set successfully");
                    }
                    event.completed();
                }
            );
        } else if (Office.context.mailbox.diagnostics.hostName !== "OutlookIOS") {
            // Fallback for older desktop clients that don't support setSignatureAsync
            logWarn("setSignatureAsync not supported — using prependAsync fallback");
            bodyItem.prependAsync(
                signatureHtml,
                { coercionType: Office.CoercionType.Html },
                function (asyncResult) {
                    clearTimeout(safetyTimeout);
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        logError("prependAsync failed: " + asyncResult.error.message);
                        notifyUser("error", "Could not set signature: " + asyncResult.error.message);
                    } else {
                        logInfo("Signature prepended via fallback");
                    }
                    event.completed();
                }
            );
        } else {
            // iOS detected — prependAsync not supported in Message Compose, exit cleanly
            logWarn("iOS detected, setSignatureAsync unavailable — skipping");
            clearTimeout(safetyTimeout);
            event.completed();
        }

    } catch (error) {
        clearTimeout(safetyTimeout);
        logError("Error: " + error.message);
        notifyUser("error", "Could not load signature. Check your connection.");
        event.completed();
    }
}

// NO CHANGE — registers the event handler with Office
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);

}); // end of Office.onReady()