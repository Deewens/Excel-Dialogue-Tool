/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console, Excel, document */

import context = Office.context;

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

/*function importCSV(event: Office.AddinCommands.Event) {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/import-csv-dialog.html",
    { height: 100, width: 100, displayInIframe: true },
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => tryCatch(processImportCsvMessage(arg)));
    }
  );

  event.completed();
}

async function processImportCsvMessage(arg) {
  await Excel.run(async (context) => {
    let editorSheet = context.workbook.worksheets.getItemOrNullObject("Dialogue Editor");

    await context.sync();

    if (editorSheet.isNullObject) {
      editorSheet = context.workbook.worksheets.add("Dialogue Editor");
    } else {
      editorSheet.getUsedRange().clear();
    }

    const headers = [
      "ID",
      "Speaker",
      "Text",
      "LocNamespace",
      "LocKey",
      "Conditions",
      "NextLine",
      "Choices",
      "SpeakerData",
    ];

    const tableHeaderRange = editorSheet.getRangeByIndexes(0, 0, 1, headers.length);

    const dialoguesTable = editorSheet.tables.add(tableHeaderRange, true);
    dialoguesTable.name = "DialogueTable";
    dialoguesTable.getHeaderRowRange().values = [headers];

    const data = JSON.parse(arg.message) as UEDialogueDataTable[];

    data.forEach((row) => {
      const FTextComponents = extractFTextComponents(row.DialogueText);

      if (FTextComponents.length > 0) FTextComponents[2] = FTextComponents[2].replace(/\\"/g, '"');

      const newData = [
        [
          row["---"],
          row.Speaker,
          FTextComponents[2],
          FTextComponents[0],
          FTextComponents[1],
          row.Conditions,
          row.NextLine,
          row.Choices,
          row.SpeakerData,
        ],
      ];

      dialoguesTable.rows.add(null, newData);
    });

    editorSheet.getUsedRange().format.autofitColumns();
    editorSheet.getUsedRange().format.autofitRows();

    editorSheet.activate();

    await context.sync();

    dialog.close();
  });
}

function extractFTextComponents(FTextString: string) {
  // Reset `lastIndex` if this regex is defined globally
  REGEX_FTEXT_EXTRACTION.lastIndex = 0;

  let regexMatches: RegExpExecArray;

  const resultArray: string[] = [];

  while ((regexMatches = REGEX_FTEXT_EXTRACTION.exec(FTextString)) !== null) {
    // This is necessary to avoid infinite loops with zero-width matches
    if (regexMatches.index === REGEX_FTEXT_EXTRACTION.lastIndex) {
      REGEX_FTEXT_EXTRACTION.lastIndex++;
    }

    regexMatches.forEach((match, groupIndex) => {
      if (groupIndex == 1) {
        resultArray.push(match);
      }
    });
  }

  return resultArray;
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

let dialog: Office.Dialog = null;*/

// Register the function with Office.
Office.actions.associate("action", action);
