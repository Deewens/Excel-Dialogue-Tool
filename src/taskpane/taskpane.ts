import { extractFTextComponents, returnObjectFromValues, parseCSV } from "../utils";
import { DialogueTable, UEDialogueDataTable } from "../types";
import Papa, { ParseError, ParseMeta } from "papaparse";
import { showOpenFilePicker, showSaveFilePicker } from "native-file-system-adapter";
import { ErrorsHandler } from "../errors-handler";

/* global Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("app-body").style.display = "block";
    document.getElementById("open-file-dialog").onclick = () => tryCatch(importCSV);
    document.getElementById("export-csv").onclick = () => tryCatch(exportCSV);
  }
});

export async function importCSV() {
  resetAndHideErrors();

  await Excel.run(async (context) => {
    const [fileHandle] = await showOpenFilePicker({
      // @ts-ignore (Type definition is out-of-date)
      types: [{ description: "CSV file", accept: { "text/csv": [".csv"] } }],
      excludeAcceptAllOption: true,
    });

    const file = await fileHandle.getFile();
    parseCSV(file, async function (data: UEDialogueDataTable[], meta: ParseMeta, error: ParseError[]) {
      globalErrors.addParseError(error);

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
        "NextLineID",
        "Choices",
        "SpeakerData",
      ];

      const tableHeaderRange = editorSheet.getRangeByIndexes(0, 0, 1, headers.length);

      const dialoguesTable = editorSheet.tables.add(tableHeaderRange, true);
      dialoguesTable.name = "DialogueTable";
      dialoguesTable.getHeaderRowRange().values = [headers];

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
            row.NextLineID,
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
    });
  });
}

export async function exportCSV() {
  resetAndHideErrors();

  await Excel.run(async (context) => {
    const editorSheet = context.workbook.worksheets.getItemOrNullObject("Dialogue Editor");

    await context.sync();

    if (editorSheet.isNullObject) {
      globalErrors.addAddinError([`Nothing to export`]);
      console.error("Nothing to export");
      displayErrorsIfExists();
      return;
    }

    const dialogueTable = editorSheet.tables.getItemOrNullObject("DialogueTable");

    await context.sync();

    if (dialogueTable.isNullObject) {
      globalErrors.addAddinError([
        `The table does not exist or has been renamed. The table should be named "DialoguesTable"`,
      ]);
      console.error(`The table does not exist or has been renamed. The table should be named "DialoguesTable"`);
      displayErrorsIfExists();
      return;
    }

    const dialogueDataRange = dialogueTable.getRange().load({ values: true });
    await context.sync();

    const valuesAsJson = returnObjectFromValues<DialogueTable>(dialogueDataRange.values);

    const formatedJson: UEDialogueDataTable[] = valuesAsJson.map((value): UEDialogueDataTable => {
      let dialogueText: string;

      value.Text = value.Text.replace(/"/g, `\\"`); // Replace the quote specific to the text, with an "escape" character that Unreal can understand

      // We need to properly reformat the localisation data. If there is no localisation data, then we just leave it empty and Unreal will generate one for us during import
      if (value.LocNamespace && value.LocKey) {
        dialogueText = `NSLOCTEXT("${value.LocNamespace}", "${value.LocKey}", "${value.Text}")`;
      } else {
        dialogueText = value.Text;
      }

      return {
        "---": value.ID,
        Speaker: value.Speaker,
        DialogueText: dialogueText,
        Conditions: value.Conditions,
        NextLineID: value.NextLineID,
        Choices: value.Choices,
        SpeakerData: value.SpeakerData,
      };
    });

    let csv = Papa.unparse(formatedJson, {
      quotes: [false, true, true, true, true, true, true],
      skipEmptyLines: "greedy",
    });

    csv += "\r\n";
    (document.getElementById("csvTextArea") as HTMLTextAreaElement).textContent = csv;

    const fileHandle = await showSaveFilePicker({
      suggestedName: "test.csv",
      types: [{ accept: { "text/csv": [".csv"] } }],
      excludeAcceptAllOption: true,
    });

    const writer = await fileHandle.createWritable();

    const blob = new Blob([csv], { type: "text/csv" });
    await writer.write(blob);
    await writer.close();
  });
}

function displayErrorsIfExists() {
  const errorAlert = document.getElementById("errorAlert");

  if (globalErrors.parseErrors.length > 0 || globalErrors.addinErrors.length > 0) {
    if (globalErrors.parseErrors.length > 0) {
      errorAlert.insertAdjacentHTML("beforeend", `<p class="fw-bold">Parsing errors:</p>`);
      errorAlert.insertAdjacentHTML("beforeend", `<ul>`);

      globalErrors.parseErrors.forEach((error) => {
        errorAlert.insertAdjacentHTML("beforeend", `<li>[${error.code}]: ${error.message}</li>`);
      });

      errorAlert.insertAdjacentHTML("beforeend", `</ul>`);
    }

    if (globalErrors.addinErrors.length > 0) {
      errorAlert.insertAdjacentHTML("beforeend", `<p class="fw-bold">Add-in errors:</p>`);
      errorAlert.insertAdjacentHTML("beforeend", `<ul>`);

      globalErrors.addinErrors.forEach((error) => {
        errorAlert.insertAdjacentHTML("beforeend", `<li>${error}</li>`);
      });

      errorAlert.insertAdjacentHTML("beforeend", `</ul>`);
    }
    errorAlert.classList.replace("d-none", "d-block");
  }
}

function resetAndHideErrors() {
  const errorAlert = document.getElementById("errorAlert");

  errorAlert.classList.replace("d-block", "d-none");
  errorAlert.innerHTML = "";

  globalErrors.clear();
}

export async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    globalErrors.addAddinError([error]);
    displayErrorsIfExists();

    console.error(error);
  }
}

let globalErrors: ErrorsHandler;
