console.log("HC Smart Alert: launchevent.js loaded");

function onMessageSendHandler(event) {
  console.log("HC Smart Alert: onMessageSendHandler fired");

  Office.context.mailbox.item.getAttachmentsAsync(function (result) {

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.log(
        "HC Smart Alert: attachment lookup failed",
        result.error
      );

      event.completed({ allowEvent: true });
      return;
    }

    var attachments = result.value || [];

    console.log(
      "HC Smart Alert: attachment count =",
      attachments.length
    );

    if (attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    var attachment = attachments[0];

    console.log("HC Smart Alert: retrieving attachment", {
      name: attachment.name,
      size: attachment.size,
      attachmentType: attachment.attachmentType
    });

    Office.context.mailbox.item.getAttachmentContentAsync(
      attachment.id,
      function (contentResult) {

        if (
          contentResult.status !==
          Office.AsyncResultStatus.Succeeded
        ) {
          console.log(
            "HC Smart Alert: content retrieval FAILED",
            contentResult.error
          );

          event.completed({
            allowEvent: false,
            errorMessage:
              "Attachment detected. Content retrieval failed. You can still choose Send anyway."
          });

          return;
        }

        var content = contentResult.value;

        console.log(
          "HC Smart Alert: content retrieval SUCCEEDED"
        );

        console.log(
          "HC Smart Alert: content format =",
          content.format
        );

        console.log(
          "HC Smart Alert: content length =",
          content.content ? content.content.length : 0
        );

        // Diagnostic only:
        // Decode the Base64 attachment and display only
        // the first 16 bytes as hexadecimal.
        if (content.content) {
          try {
            var binary = atob(content.content);

            var signature = [];

            for (
              var i = 0;
              i < Math.min(16, binary.length);
              i++
            ) {
              signature.push(
                binary
                  .charCodeAt(i)
                  .toString(16)
                  .padStart(2, "0")
              );
            }

            console.log(
              "HC Smart Alert: first 16 bytes =",
              signature.join(" ")
            );

          } catch (e) {
            console.log(
              "HC Smart Alert: Base64 decode failed =",
              e.message
            );
          }
        }

        event.completed({
          allowEvent: false,
          errorMessage:
            "Attachment content was successfully retrieved. You can still choose Send anyway."
        });
      }
    );
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
