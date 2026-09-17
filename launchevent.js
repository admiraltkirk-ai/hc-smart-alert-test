console.log("HC Smart Alert: launchevent.js loaded");

function onMessageSendHandler(event) {
  console.log("HC Smart Alert: onMessageSendHandler fired");

  Office.context.mailbox.item.getAttachmentsAsync(function (result) {

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("HC Smart Alert: attachment lookup failed");
      event.completed({ allowEvent: true });
      return;
    }

    var attachments = result.value || [];

    console.log(
      "HC Smart Alert: attachment count = " + attachments.length
    );

    if (attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    event.completed({
      allowEvent: false,
      errorMessage:
        "Attachment test: this message contains an attachment. You can still choose Send anyway."
    });
  });
}

Office.onReady(function () {
  console.log("HC Smart Alert: Office.js ready");

  Office.actions.associate(
    "onMessageSendHandler",
    onMessageSendHandler
  );

  console.log("HC Smart Alert: handler associated");
});
