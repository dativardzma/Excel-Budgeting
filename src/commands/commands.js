// /*
// //  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// //  * See LICENSE in the project root for license information.
// //  */

// // /* global Office */

// // Office.onReady(() => {
// //   // If needed, Office.js is ready to be called.
// // });

// // /**
// //  * Shows a notification when the add-in command is executed.
// //  * @param event {Office.AddinCommands.Event}
// //  */
// // function action(event) {
// //   const message = {
// //     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
// //     message: "Performed action.",
// //     icon: "Icon.80x80",
// //     persistent: true,
// //   };

// //   // Show a notification message.
// //   Office.context.mailbox.item?.notificationMessages.replaceAsync(
// //     "ActionPerformanceNotification",
// //     message
// //   );

// //   // Be sure to indicate when the add-in command function is complete.
// //   event.completed();
// // }
// // function insertMaterials(event) {
// //   Excel.run(function (context) {
// //     const sheet = context.workbook.worksheets.getActiveWorksheet();
// //     const range = sheet.getRange("A1");
// //     range.values = [["Material", "Qty", "Price"]];
// //     return context.sync();
// //   }).then(function () {
// //     event.completed();
// //   });
// // }
// // // Register the function with Office.
// // if (typeof Office !== "undefined") {
// //   Office.actions.associate("insertMaterials", insertMaterials);
// // }

// // Office.actions.associate("action", action);


// /* global Office, OfficeRuntime, console */

// // This is the command that runs when you click "Addin projects"
// // Called when the commands runtime is ready


// /* global Office, OfficeRuntime, Excel, console */

// // This runs when you click the ribbon button "Addin projects"
// /* global Office, OfficeRuntime, Excel, console, window */

// // This runs when you click the ribbon button "Addin projects"
// /* global Office, OfficeRuntime, Excel, console, window */

// // This function is called when the ribbon button "Addin projects" is clicked
async function Hello(event) {
  try {
    // 1) Check login state stored by taskpane.js
    let loggedIn = false;
    try {
      const stored = await OfficeRuntime.storage.getItem("userLoggedIn");
      loggedIn = stored === "true";
      console.log("Hello: stored login =", stored);
    } catch (e) {
      console.warn("Hello: could not read login state:", e);
    }

    // 2) If NOT logged in -> show popup + status bar message
    if (!loggedIn) {
      // Popup so the user sees clearly what to do
      window.alert("First login.\nGo to the 'Budgetings' tab and click the 'Login' button.");

      // Also set status bar text (bottom of Excel)
      try {
        await Excel.run(async (context) => {
          context.workbook.application.statusBar =
            "First login: Budgetings → Login.";
          await context.sync();
        });
      } catch (e) {
        console.error("Hello: failed to set status bar message:", e);
      }

      event.completed();
      return;
    }

    // 3) If logged in -> do your real work
    console.log("Hello: user is logged in. Running Insert Materials logic...");

    // Example insert logic (you can customize later):
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:C1");
      range.values = [["Material", "Qty", "Price"]];
      await context.sync();
    });

  } catch (err) {
    console.error("Hello: error:", err);
  } finally {
    // IMPORTANT: always complete the event
    event.completed();
  }
}

// Called when the commands runtime is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Commands runtime ready");
  }
});

// Register function so it matches FunctionName="Hello" from manifest
if (Office.actions) {
  Office.actions.associate("Hello", Hello);
}

/* global Office, OfficeRuntime, Excel, console */

/* global Office, OfficeRuntime, Excel, console, window */

// Called when user clicks the "Addin projects" button
