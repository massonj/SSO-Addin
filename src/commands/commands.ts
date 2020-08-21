/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, console, Office, Excel, self, window */

Office.onReady(() => {
  return Excel.run((context) => {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
      .then(function () {
          console.log("Event handler successfully registered for onChanged event in the worksheet.");
      });
  });
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
export function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "action",
    message
  );

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

export function handleChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch((error) => { console.log(`caught error :${error}`)});
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

const g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.action;
