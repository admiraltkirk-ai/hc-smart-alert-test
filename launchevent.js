console.log("HC Smart Alert: label diagnostic loaded");

function onMessageSendHandler(event) {
  console.log("HC Smart Alert: handler fired");

  Office.context.mailbox.item.getAttachmentsAsync(function (result) {

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("HC Smart Alert: attachment lookup FAILED", result.error);
      event.completed({ allowEvent: true });
      return;
    }

    var attachments = result.value || [];

    console.log("HC Smart Alert: attachment count =", attachments.length);

    if (attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    var attachment = attachments[0];

    Office.context.mailbox.item.getAttachmentContentAsync(
      attachment.id,
      function (contentResult) {

        if (contentResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.log(
            "HC Smart Alert: content retrieval FAILED",
            contentResult.error
          );

          event.completed({
            allowEvent: false,
            errorMessage:
              "Attachment could not be inspected. You can still choose Send anyway."
          });

          return;
        }

        try {
          var binary = atob(contentResult.value.content);

          console.log(
            "HC Smart Alert: inspecting",
            attachment.name,
            binary.length,
            "bytes"
          );

          /*
           * Extract readable ASCII strings.
           */
          var strings = [];
          var current = "";

          for (var i = 0; i < binary.length; i++) {
            var code = binary.charCodeAt(i);

            if (code >= 32 && code <= 126) {
              current += String.fromCharCode(code);
            } else {
              if (current.length >= 4) {
                strings.push(current);
              }

              current = "";
            }
          }

          if (current.length >= 4) {
            strings.push(current);
          }

          /*
           * Find the MSO rights-label metadata.
           */
          var labelBlocks = strings.filter(function (text) {
            var lower = text.toLowerCase();

            return (
              lower.indexOf("mso:soft rights label") !== -1 ||
              lower.indexOf("name=lcid") !== -1
            );
          });

          console.log(
            "HC Smart Alert: rights label blocks =",
            labelBlocks.length
          );

          labelBlocks.forEach(function (block, index) {

            console.log(
              "========== HC LABEL BLOCK " +
              (index + 1) +
              " =========="
            );

            console.log(block);

            console.log(
              "========== END HC LABEL BLOCK =========="
            );
          });

          /*
           * Search the label block for GUID-looking values.
           */
          var combined = labelBlocks.join(" ");

          var guidRegex =
            /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g;

          var guids = combined.match(guidRegex) || [];

          var uniqueGuids = [];

          guids.forEach(function (guid) {
            var normalised = guid.toLowerCase();

            if (uniqueGuids.indexOf(normalised) === -1) {
              uniqueGuids.push(normalised);
            }
          });

          console.log(
            "HC Smart Alert: GUID candidates =",
            uniqueGuids
          );

          /*
           * Diagnostic popup only.
           */
          event.completed({
            allowEvent: false,
            errorMessage:
              "Sensitivity-label metadata was inspected. You can still choose Send anyway."
          });

        } catch (e) {

          console.log(
            "HC Smart Alert: label diagnostic FAILED",
            e
          );

          event.completed({
            allowEvent: false,
            errorMessage:
              "Label diagnostic failed. You can still choose Send anyway."
          });
        }
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
