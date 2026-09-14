function onMessageSendHandler(event) {
  Office.context.mailbox.item.getAttachmentsAsync(
    { asyncContext: event },
    function (asyncResult) {
      var sendEvent = asyncResult.asyncContext;
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        sendEvent.completed({ allowEvent: true });
        return;
      }
      var attachments = asyncResult.value || [];
      if (attachments.length === 0) {
        sendEvent.completed({ allowEvent: true });
        return;
      }
      sendEvent.completed({
        allowEvent: false,
        errorMessage: "Attachment test: this message contains an attachment. You can still choose Send anyway.",
        errorMessageMarkdown: "**Attachment test**\\n\\nThis message contains an attachment. You can still choose **Send anyway**."
      });
    }
  );
}
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
