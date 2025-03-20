/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console, Excel, document */

import Worksheet = Excel.Worksheet;
import RequestContext = Excel.RequestContext;
import { DialogueTable, UEDialogueDataTable } from "../types";
import { ParseResult } from "papaparse";
import { config } from "../config";
import { extractFTextComponents, returnObjectFromValues } from "../utils";

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

// Register the function with Office.
Office.actions.associate("action", action);

let importDialog: Office.Dialog = null;
let exportDialog: Office.Dialog = null;

async function onImportCSVClicked(event) {
  try {
    await Excel.run(async (context) => {
      Office.context.ui.displayDialogAsync(
        // "https://localhost:3000/import-csv-dialog.html",
        "https://sadspoonstorage.z6.web.core.windows.net/import-csv-dialog.html",
        {
          height: 25,
          width: 35,
          displayInIframe: true,
        },
        (asyncResult) => {
          importDialog = asyncResult.value;
          importDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
        }
      );

      event.completed();
    });
  } catch (error) {
    console.error(error);
  }

  async function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message) as ParseResult<UEDialogueDataTable>;

    if (messageFromDialog.errors.length === 0) {
      importDialog.close();

      await createOrUpdateExcel(messageFromDialog.data);
    }
  }
}

Office.actions.associate("importCSV", onImportCSVClicked);

async function onExportCSVClicked(event) {
  try {
    await Excel.run(async (context) => {
      const dialogueDataJson = await getCSVDataToExport();

      Office.context.ui.displayDialogAsync(
        // "https://localhost:3000/export-csv-dialog.html",
        "https://sadspoonstorage.z6.web.core.windows.net/export-csv-dialog.html",
        {
          height: 45,
          width: 45,
          displayInIframe: true,
        },
        (asyncResult) => {
          exportDialog = asyncResult.value;

          exportDialog.messageChild(JSON.stringify(dialogueDataJson), { targetOrigin: "https://localhost:3000" });

          // wait for the dialog to tell us its ready to process messages because microsoft is shit. Once it is, we can sent messages to it
          exportDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
            if (arg.message === "IAmReady") {
              console.log(dialogueDataJson);
              // exportDialog.messageChild(JSON.stringify(dialogueDataJson), { targetOrigin: "https://localhost:3000" });
            }
          });
        }
      );

      event.completed();
    });
  } catch (error) {
    console.error(error);
  }
}

Office.actions.associate("exportCSV", onExportCSVClicked);

async function getCSVDataToExport() {
  try {
    return await Excel.run(async (context) => {
      const editorSheet = context.workbook.worksheets.getItemOrNullObject("Dialogue Editor");

      await context.sync();

      if (editorSheet.isNullObject) return;

      const dialogueTable = editorSheet.tables.getItemOrNullObject("DialogueTable");

      await context.sync();

      if (dialogueTable.isNullObject) return;

      const dialogueDataRange = dialogueTable.getRange().load({ values: true });
      await context.sync();

      return returnObjectFromValues<DialogueTable>(dialogueDataRange.values);
    });
  } catch (error) {
    console.error(error);
  }
}

async function selectOrCreateWorksheet(context: RequestContext, worksheetName: string) {
  let sheet = context.workbook.worksheets.getItemOrNullObject(worksheetName);

  await context.sync();

  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(worksheetName);
  }

  return sheet;
}

type CreateTableMeta = {
  address: Excel.Range;
  hasHeaders: boolean;
  headerRowValues: string[][];
};

function createTable(worksheet: Worksheet, tableName: string, meta: CreateTableMeta) {
  const dialoguesTable = worksheet.tables.add(meta.address, meta.hasHeaders);
  dialoguesTable.name = tableName;
  dialoguesTable.getHeaderRowRange().values = meta.headerRowValues;

  return dialoguesTable;
}

async function createOrUpdateExcel(data: UEDialogueDataTable[]) {
  try {
    await Excel.run(async (context) => {
      const editorSheet = await selectOrCreateWorksheet(context, config.worksheetName);

      await context.sync();

      let dialogueTable = editorSheet.tables.getItemOrNullObject(config.tableName);

      await context.sync();

      if (dialogueTable.isNullObject) {
        const tableHeaderRange = editorSheet.getRangeByIndexes(0, 0, 1, config.tableHeaderRows.length);
        tableHeaderRange.setRowProperties([{}]);
        dialogueTable = createTable(editorSheet, config.tableName, {
          address: tableHeaderRange,
          hasHeaders: true,
          headerRowValues: [config.tableHeaderRows],
        });
      } else {
        const rowsCount = dialogueTable.rows.getCount();

        await context.sync();

        // eslint-disable-next-line office-addins/load-object-before-read
        dialogueTable.rows.deleteRowsAt(0, rowsCount.value - 1);
      }

      data.forEach((row) => {
        const FTextComponents = extractFTextComponents(row.DialogueText);

        if (FTextComponents.length > 0) FTextComponents[2] = FTextComponents[2].replace(/\\"/g, '"');

        const text = FTextComponents[2];
        const locNamespace = FTextComponents[0];
        const locKey = FTextComponents[1];

        const newData = [
          [
            row["---"],
            row.Speaker,
            text,
            row.NextLineID,
            row.Choices,
            row.Conditions,
            locNamespace,
            locKey,
            row.SpeakerData,
          ],
        ];

        dialogueTable.rows.add(null, newData);
      });

      editorSheet.getUsedRange().format.autofitColumns();
      editorSheet.getUsedRange().format.autofitRows();

      editorSheet.activate();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
