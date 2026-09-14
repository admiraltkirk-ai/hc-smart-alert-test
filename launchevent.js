console.log("HC Smart Alert: launchevent.js loaded");

function onMessageSendHandler(event) {
  console.log("HC Smart Alert: onMessageSendHandler fired");

  event.completed({ allowEvent: true });
}

Office.actions.associate(
  "onMessageSendHandler",
  onMessageSendHandler
);
