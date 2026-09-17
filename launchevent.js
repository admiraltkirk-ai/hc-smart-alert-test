console.log("HC Smart Alert: metadata diagnostic loaded");

function onMessageSendHandler(event) {
  console.log("HC Smart Alert: onMessageSendHandler fired");

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

    console.log("HC Smart Alert: examining attachment", {
      name: attachment.name,
      size: attachment.size,
      attachmentType: attachment.attachmentType
    });

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
              "Attachment detected, but its content could not be examined. You can still choose Send anyway."
          });
          return;
        }

        var content = contentResult.value;

        console.log("HC Smart Alert: content retrieval SUCCEEDED");
        console.log("HC Smart Alert: content format =", content.format);

        try {
          var binary = atob(content.content);

          console.log("HC Smart Alert: decoded bytes =", binary.length);

          // Confirm the OLE / Compound File signature.
          var signature = [];

          for (var i = 0; i < Math.min(8, binary.length); i++) {
            signature.push(
              binary.charCodeAt(i).toString(16).padStart(2, "0")
            );
          }

          console.log(
            "HC Smart Alert: file signature =",
            signature.join(" ")
          );

          /*
           * Extract runs of readable ASCII characters.
           * We deliberately do NOT print the whole document.
           */
          var readableStrings = [];
          var current = "";

          for (var j = 0; j < binary.length; j++) {
            var code = binary.charCodeAt(j);

            if (code >= 32 && code <= 126) {
              current += String.fromCharCode(code);
            } else {
              if (current.length >= 4) {
                readableStrings.push(current);
              }
              current = "";
            }
          }

          if (current.length >= 4) {
            readableStrings.push(current);
          }

          var keywords = [
            "label",
            "msip",
            "sensitivity",
            "encrypted",
            "encryption",
            "microsoft",
            "protection",
            "protected",
            "publishing",
            "license",
            "irm",
            "drm"
          ];

          var matches = [];

          readableStrings.forEach(function (text) {
            var lower = text.toLowerCase();

            var matched = keywords.some(function (keyword) {
              return lower.indexOf(keyword) !== -1;
            });

            if (matched) {
              // Avoid huge console output.
              var safeText =
                text.length > 300
                  ? text.substring(0, 300) + "..."
                  : text;

              if (matches.indexOf(safeText) === -1) {
                matches.push(safeText);
              }
            }
          });

          console.log(
            "HC Smart Alert: metadata keyword matches =",
            matches.length
          );

          matches.slice(0, 30).forEach(function (match, index) {
            console.log(
              "HC Smart Alert: MATCH " + (index + 1) + " =",
              match
            );
          });

          if (matches.length === 0) {
            console.log(
              "HC Smart Alert: no readable protection metadata found"
            );
          }

          event.completed({
            allowEvent: false,
            errorMessage:
              "HC attachment diagnostic completed. You can still choose Send anyway."
          });

        } catch (e) {
          console.log(
            "HC Smart Alert: diagnostic FAILED =",
            e
          );

          event.completed({
            allowEvent: false,
            errorMessage:
              "Attachment diagnostic failed. You can still choose Send anyway."
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
