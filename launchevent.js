/* HC Smart Alert proof-of-concept.
   This first test deliberately warns on every send.
   It does not yet inspect recipients, labels, or attachments. */

function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "HC Smart Alert test: this is a warning only. You can still choose Send Anyway.",
    errorMessageMarkdown: "**HC Smart Alert test**\n\nThis is a warning only. You can still choose **Send Anyway**."
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
