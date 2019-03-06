var config;
var btnEvent;

// The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {
  config = getConfig();
};

// Add any ui-less function here
function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('error', {
    type: 'errorMessage',
    message: error
  }, function(result){
  });
}

var settingsDialog;


function receiveMessage(message) {
  config = JSON.parse(message.message);
  setConfig(config, function(result) {
    settingsDialog.close();
    settingsDialog = null;
    btnEvent.completed();
    btnEvent = null;
  });
}

function dialogClosed(message) {
  settingsDialog = null;
  btnEvent.completed();
  btnEvent = null;
}

