/* global Office, console */
/*
 * Connacht Hospitality Group — Outlook Signature Add-in
 * autorun.js — Fetches employee data and sets the signature
 */

var API_URL = "https://connachtsignatures-bsbfakbbcjf6fnbb.westeurope-01.azurewebsites.net/api/signature";

var HOTEL_CONFIG = {
    "@chgl.ie": {
        website: "www.chgl.ie",
        address: "Connacht Hotel, Old Dublin Rd, Galway, H91 K5DD",
        websiteUrl: "https://www.chgl.ie/",
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

// ── Logging Helpers ───────────────────────────────────────
var LOG = "[ConnachtSig]";
function logInfo(msg)  { console.log(LOG, "INFO:", msg); }
function logWarn(msg)  { console.warn(LOG, "WARN:", msg); }
function logError(msg) { console.error(LOG, "ERROR:", msg); }

// ── Notification Helper ───────────────────────────────────
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

// ── Hotel Config Lookup ───────────────────────────────────
function getConfigForEmail(email) {
    var emailLower = email.toLowerCase();
    for (var suffix in HOTEL_CONFIG) {
        if (suffix !== "default" && emailLower.endsWith(suffix)) {
            return HOTEL_CONFIG[suffix];
        }
    }
    return HOTEL_CONFIG["default"];
}

// ── Fetch Employee Data from Azure Function ───────────────
async function getEmployeeData() {
    var userEmail = Office.context.mailbox.userProfile.emailAddress;
    logInfo("Current user email: " + userEmail);

    var response = await fetch(API_URL + "?email=" + encodeURIComponent(userEmail));

    if (response.status === 404) {
        logWarn("No employee found in Azure AD for: " + userEmail);
        return null;
    }

    if (!response.ok) {
        throw new Error("API request failed — status " + response.status);
    }

    var employee = await response.json();
    return employee;
}

// ── Build Signature HTML ──────────────────────────────────
function buildSignatureHtml(emp, config) {
    var s = config.style;

    var html = ''
        + '<table cellpadding="0" cellspacing="0" border="0" style="font-family:' + s.fontFamily + ';font-size:12px;color:' + s.textColor + ';line-height:1.5;">'
        + '<tr>'

        + '<td style="padding-right:20px;vertical-align:top;">'
        + (emp.name ? '<strong style="font-size:' + s.nameSize + ';color:' + s.nameColor + ';">' + emp.name + '</strong><br/>' : '')
        + (emp.title ? '<span style="font-size:12px;color:' + s.titleColor + ';">' + emp.title + '</span>' : '')
        + '</td>'

        + '<td style="padding-left:20px;vertical-align:top;border-left:1px solid ' + s.dividerColor + ';">'
        + (emp.email ? '<span style="padding-left:10px;"><strong>E:</strong> <a href="mailto:' + emp.email + '" style="color:' + s.linkColor + ';text-decoration:underline;">' + emp.email + '</a></span><br/>' : '')
        + (emp.phone ? '<span style="padding-left:10px;"><strong>T:</strong> <a href="tel:' + emp.phone + '" style="color:' + s.linkColor + ';text-decoration:underline;">' + emp.phone + '</a></span><br/>' : '')
        + (emp.website ? '<span style="padding-left:10px;"><strong>W:</strong> <a href="https://' + emp.website + '" target="_blank" style="color:' + s.linkColor + ';text-decoration:underline;">' + emp.website + '</a></span><br/>' : '')
        + (emp.address ? '<span style="padding-left:10px;"><strong>A:</strong> ' + emp.address + '</span>' : '')
        + '</td>'

        + '</tr>'
        + '</table>'

        + '<table cellpadding="0" cellspacing="0" border="0" style="padding-top:15px;">'
        + '<tr>'
        + '<td>'
        + '<a href="' + config.websiteUrl + '" target="_blank" style="text-decoration:none;">'
        + (emp.banner ? '<img src="' + emp.banner + '" alt="Connacht Hospitality Group" width="500" style="border:0;display:block;" />' : '')
        + '</a>'
        + '</td>'
        + '</tr>'
        + '</table>'

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

// ── Main LaunchEvent Handler ──────────────────────────────
// IMPORTANT: This function and Office.actions.associate must be outside
// Office.onReady() so iOS Outlook can call them immediately on file load,
// before the onReady callback has fired.
async function onNewMessageCompose(event) {
    logInfo("OnNewMessageCompose triggered");

    // Safety timeout — mobile kills the add-in after 60s, this exits cleanly at 55s
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

        var config = getConfigForEmail(employee.email);
        employee.website = employee.website || config.website;
        employee.address = employee.address || config.address;

        logInfo("Config applied — website: " + employee.website);

        var signatureHtml = buildSignatureHtml(employee, config);

        if (bodyItem.setSignatureAsync) {
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
        } else {
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
        }

    } catch (error) {
        clearTimeout(safetyTimeout);
        logError("Error: " + error.message);
        notifyUser("error", "Could not load signature. Check your connection.");
        event.completed();
    }
}

// Registers the handler — must also be outside Office.onReady()
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);