function onMessageSendHandler(event) {
  Office.context.mailbox.item.getAttachmentsAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var attachments = result.value || [];
      if (attachments.length === 0) {
        event.completed({ allowEvent: true });
        return;
      }
      event.completed({
        allowEvent: false,
        errorMessage: "Attachment test: this message contains an attachment. You can still choose Send anyway.",
        errorMessageMarkdown: "**Attachment test**\n\nThis message contains an attachment. You can still choose **Send anyway**."
      });
      return;
    }
    event.completed({ allowEvent: true });
  });
}
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

