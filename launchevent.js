console.log("HC Smart Alert: GUID-LABEL diagnostic v6 loaded");
console.log("HC Smart Alert: raw metadata diagnostic loaded");

function onMessageSendHandler(event) {
  console.log("HC Smart Alert: handler fired");

  Office.context.mailbox.item.getAttachmentsAsync(function (result) {

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.log(
        "HC Smart Alert: attachment lookup FAILED",
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

    console.log(
      "HC Smart Alert: examining attachment =",
      attachment.name
    );

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
              "Attachment could not be inspected. You can still choose Send anyway."
          });

          return;
        }

        try {
          var binary = atob(contentResult.value.content);

          console.log(
            "HC Smart Alert: decoded bytes =",
            binary.length
          );

          /*
           * Convert non-printable bytes to spaces while
           * preserving the original byte positions.
           */
          var searchable = "";

          for (var i = 0; i < binary.length; i++) {
            var code = binary.charCodeAt(i);

            if (
              (code >= 32 && code <= 126) ||
              code === 9 ||
              code === 10 ||
              code === 13
            ) {
              searchable += binary.charAt(i);
            } else {
              searchable += " ";
            }
          }

          /*
           * Search for markers observed inside the
           * protected Office container.
           */
          var markers = [
            "MSO:soft Rights Label",
            "Rights Label",
            "Highly Confidential",
            "NAME=LCID",
            "MS-DRM-Server",
            "ISSUEDTIME",
            "DESCRIPTOR"
          ];

          var foundAny = false;

          markers.forEach(function (marker) {

            var position = searchable
              .toLowerCase()
              .indexOf(marker.toLowerCase());

            console.log(
              'HC Smart Alert: marker "' +
                marker +
                '" position =',
              position
            );

            if (position !== -1) {
              foundAny = true;

              var start = Math.max(0, position - 500);
              var end = Math.min(
                searchable.length,
                position + 1500
              );

              var section = searchable.substring(
                start,
                end
              );

              console.log(
                "========== RAW METADATA AROUND " +
                  marker +
                  " =========="
              );

              console.log(section);

              console.log(
                "========== END RAW METADATA =========="
              );
            }
          });

          /*
           * Collect GUID-shaped values from the
           * protected document metadata.
           */
          var guidRegex =
            /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g;

          var guidMatches =
            searchable.match(guidRegex) || [];

          var uniqueGuids = [];

          guidMatches.forEach(function (guid) {
            var normalised = guid.toLowerCase();

            if (
              uniqueGuids.indexOf(normalised) === -1
            ) {
              uniqueGuids.push(normalised);
            }
          });

          console.log(
            "HC Smart Alert: ALL GUID candidates =",
            uniqueGuids
          );

          /*
           * Find the Highly Confidential FINANCE label
           * and inspect the metadata immediately before it.
           */
          var labelName =
            "Highly Confidential FINANCE";

          var labelPosition = searchable
            .toLowerCase()
            .indexOf(labelName.toLowerCase());

          console.log(
            "HC Smart Alert: FINANCE label position =",
            labelPosition
          );

          if (labelPosition !== -1) {

            var labelStart = Math.max(
              0,
              labelPosition - 1000
            );

            var labelSection =
              searchable.substring(
                labelStart,
                labelPosition + 300
              );

            var labelGuids =
              labelSection.match(guidRegex) || [];

            console.log(
              "HC Smart Alert: GUIDs immediately before FINANCE label =",
              labelGuids
            );

            console.log(
              "HC Smart Alert: FINANCE label metadata block =",
              labelSection
            );
          } else {
            console.log(
              "HC Smart Alert: FINANCE label NOT FOUND"
            );
          }

          console.log(
            "HC Smart Alert: metadata marker found =",
            foundAny
          );

          event.completed({
            allowEvent: false,
            errorMessage:
              "Protected-document metadata diagnostic completed. You can still choose Send anyway."
          });

        } catch (e) {

          console.log(
            "HC Smart Alert: raw diagnostic FAILED",
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

  console.log(
    "HC Smart Alert: Office.js ready"
  );

  Office.actions.associate(
    "onMessageSendHandler",
    onMessageSendHandler
  );

  console.log(
    "HC Smart Alert: handler associated"
  );
});
