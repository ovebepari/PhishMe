/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

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

function sucessNotif() {
  var id = "0";
  var details = {
    type: "informationalMessage",
    icon: "Icon.16x16",
    message: "Email forward successful!",
    persistent: false
  };
  Office.context.mailbox.item.notificationMessages.addAsync(id, details, function(value) {});
}

function failedNOtif() {
  var id = "0";
  var details = {
    type: "informationalMessage",
    icon: "Icon.16x16",
    message: "Email forward failed!",
    persistent: false
  };
  Office.context.mailbox.item.notificationMessages.addAsync(id, details, function(value) {});
}

function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === "OutlookIOS") {
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

function simpleForwardEmail() {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
    var ewsId = Office.context.mailbox.item.itemId;
    var accessToken = result.value;
    simpleForwardFunc(accessToken);
  });
}

simpleForwardEmail();

function simpleForwardFunc(accessToken) {
  var itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  var forwardUrl = Office.context.mailbox.restUrl + "/v1.0/me/messages/" + itemId + "/forward";

  const metadata = JSON.stringify({
    Comment: "FYI",
    ToRecipients: [
      {
        EmailAddress: {
          Name: "Ove Bepari",
          Address: "ovebepari@gmail.com"
        }
      }
    ]
  });

  var response = $.ajax({
    url: forwardUrl,
    type: "POST",
    dataType: "json",
    contentType: "application/json",
    data: metadata,
    headers: { Authorization: "Bearer " + accessToken }
  }).always(function(response){
    if(response.status.toString() == '202'){
      sucessNotif();
    }
    else{
      failedNOtif();
    }
  });
}
