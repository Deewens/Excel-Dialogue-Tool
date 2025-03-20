import { DialogueTable, UEDialogueDataTable } from "../types";
import { ErrorsHandler } from "../errors-handler";
import Papa from "papaparse";
import { showSaveFilePicker } from "native-file-system-adapter";

/* global Office, Excel */

Office.onReady((info) => {
  Office.context.ui.messageParent("IAmReady");

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    onMessageFromParent,
    onRegisterMessageComplete
  );
  document.getElementById("exportCsvButton").onclick = () => tryCatch(exportCSV);
});

// TODO: remove the ugly code duplication but i'm lazy, it's late, and I already spent too much time on this, for now it worksm

let globalErrors = new ErrorsHandler();
let dataReceivedFromParent = false;
let dialogueData: DialogueTable[] = [];

function onMessageFromParent(arg) {
  try {
    dialogueData = JSON.parse(arg.message);
  } catch (e) {
    console.error("JSON Parsing error: " + e);
  }

  const exportCsvButton = document.getElementById("exportCsvButton") as HTMLButtonElement;

  exportCsvButton.disabled = false;

  dataReceivedFromParent = true;

  const formatedJson: UEDialogueDataTable[] = dialogueData.map((value): UEDialogueDataTable => {
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
}

function onRegisterMessageComplete(asyncResult) {
  console.log(asyncResult.status);
  if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
    console.error(asyncResult.error.message);
  }
}

async function exportCSV() {
  resetAndHideErrors();

  if (dataReceivedFromParent == false) return;

  const formatedJson: UEDialogueDataTable[] = dialogueData.map((value): UEDialogueDataTable => {
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

  const fileHandle = await showSaveFilePicker({
    suggestedName: "test.csv",
    types: [{ accept: { "text/csv": [".csv"] } }],
    excludeAcceptAllOption: true,
  });

  const writer = await fileHandle.createWritable();

  const blob = new Blob([csv], { type: "text/csv" });
  await writer.write(blob);
  await writer.close();
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
