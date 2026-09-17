console.log("HC Smart Alert: launchevent.js loaded");

function onMessageSendHandler(event) {
  console.log("HC Smart Alert: onMessageSendHandler fired");
  event.completed({ allowEvent: true });
}

Office.onReady(function () {
  console.log("HC Smart Alert: Office.js ready");

  Office.actions.associate(
    "onMessageSendHandler",
    onMessageSendHandler
  );

  console.log("HC Smart Alert: handler associated");
});
