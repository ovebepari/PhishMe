/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}


function success(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Email Forward Successful!",
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function failed(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message: "Email Forwarding failed, contact the add in developer!",
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}


function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
g.success = success;
g.failed = failed;
g.forwardEmail = forwardEmail;



// Ove starts to code from here
// code found in: https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/use-rest-api
function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}


// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;


// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/master/samples/outlook/85-tokens-and-service-calls/basic-rest-cors.yaml
function forwardEmail(event){
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
      var ewsId = Office.context.mailbox.item.itemId;
      var token = result.value;

      (function(accessToken) {
        // Get the item's REST ID.
        var itemId = getItemRestId();
        const forward = {toRecipients:[{emailAddress:{address:"ovebepari@gmail.com"}}]};

        // Construct the REST URL to the current item.
        // Details for formatting the URL can be found at
        // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
        var forwardUrl = Office.context.mailbox.restUrl +
          '/v2.0/me/messages/' + itemId + '/createForward';

        $.ajax({
          type: "POST",
          url: forwardUrl,
          dataType: 'json',
          data: forward,
          headers: { 'Authorization': 'Bearer ' + accessToken }
        }).done(function(item){
          success();
        }).fail(function(error){
          failed();
        });
      })(token);
  });
};
