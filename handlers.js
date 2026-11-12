/*
 * Outlook recipient-specific introduction
 * Intended for Outlook Classic / Outlook on the web / new Outlook with
 * Office.js OnMessageRecipientsChanged event support (Mailbox 1.11).
 *
 * CONFIGURATION: Edit the list below. Match only recipients in the To field.
 * Use \n for a line break within the text string.
 */
const RECIPIENT_RULES = [
  {
    address: "halabsa@igo-sy.com",
    introduction: "/nتحية طيبة ،السيدة هبة المحترمة"
  },
  {
    address: "akojak@igo-sy.com",
    introduction: "/nتحية طيبة ،السيدة أماني المحترمة"
  }
];

function htmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function completeEvent(event) {
  try { event.completed(); } catch (e) { console.error("Event completion failed", e); }
}

function onRecipientsChanged(event) {
  try {
    const message = Office.context.mailbox.item;
    if (!message || !message.to || !message.body) {
      completeEvent(event);
      return;
    }

    // This event also triggers when Cc/Bcc changes; To is intentionally
    // re-evaluated each time, which is harmless due to the duplicate check.
    message.to.getAsync(function (toResult) {
      if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to read To recipients", toResult.error);
        completeEvent(event);
        return;
      }

      const addresses = toResult.value.map(function (recipient) {
        return (recipient.emailAddress || "").trim().toLowerCase();
      });
      const rule = RECIPIENT_RULES.find(function (candidate) {
        return addresses.includes(candidate.address.trim().toLowerCase());
      });
      if (!rule) {
        completeEvent(event);
        return;
      }

      // Do not insert the same intro twice if the user changes recipients
      // again or the event is fired more than once.
      message.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
        if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to read body", bodyResult.error);
          completeEvent(event);
          return;
        }
        const bodyText = (bodyResult.value || "").replace(/\r\n?/g, "\n");
        const introText = rule.introduction.replace(/\r\n?/g, "\n");
        if (bodyText.includes(introText)) {
          completeEvent(event);
          return;
        }

        message.body.getTypeAsync(function (formatResult) {
          if (formatResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to read body format", formatResult.error);
            completeEvent(event);
            return;
          }
          const isHtml = formatResult.value === Office.CoercionType.Html;
          const insertion = isHtml
            ? "<div>" + htmlEscape(rule.introduction).replace(/\n/g, "<br>") + "</div><p>&nbsp;</p>"
            : rule.introduction + "\r\n\r\n";

          message.body.prependAsync(
            insertion,
            { coercionType: formatResult.value },
            function (insertResult) {
              if (insertResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to prepend recipient introduction", insertResult.error);
              }
              completeEvent(event);
            }
          );
        });
      });
    });
  } catch (err) {
    console.error("Recipient-intro add-in error", err);
    completeEvent(event);
  }
}

// Required for classic Outlook's event-based JavaScript runtime.
Office.actions.associate("onRecipientsChanged", onRecipientsChanged);
