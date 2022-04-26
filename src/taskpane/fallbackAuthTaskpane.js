let loginDialog;

function dialogFallback() {
  const url = "/dialog.html";
  showLoginPopup(url);
}

function processMessage(arg) {
  console.log("Message received in processMessage: " + JSON.stringify(arg));
  let messageFromDialog = JSON.parse(arg.message);

  if (messageFromDialog.status === "success") {
    loginDialog.close();
    console.log("!!!!!Access token!!!!! -----> " + messageFromDialog.result);
  } else {
    loginDialog.close();
    console.error("!!!!!Error in auth!!!!! -----> " + error.toString());
  }
}

function showLoginPopup(url) {
  const fullUrl = location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + url;

  // height and width are percentages of the size of the parent Office application - Outlook
  Office.context.ui.displayDialogAsync(fullUrl, { height: 60, width: 30 }, function (result) {
    console.log("Dialog has initialized. Wiring up events");
    loginDialog = result.value;
    console.log("!!!!!result!!!!!" + result.toString());
    console.log("!!!!!loginDialog!!!!!" + loginDialog.toString());
    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  });
}

module.exports = dialogFallback;
