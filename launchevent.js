function onMessageSendHandler(event) {
  event.completed({ allowEvent: true });
}

Office.onReady(function () {
  console.log("Office.js ready - registering Smart Alert handler");

  Office.actions.associate(
    "onMessageSendHandler",
    onMessageSendHandler
  );
});
