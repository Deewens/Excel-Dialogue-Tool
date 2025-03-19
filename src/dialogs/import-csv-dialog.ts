/* global Office, Excel */

import { extractFTextComponents, parseCSV } from "../utils";
import { ImportDialogMessage, UEDialogueDataTable } from "../types";
import Papa, { ParseError, ParseMeta } from "papaparse";
import { ErrorsHandler } from "../errors-handler";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("fileInput").onchange = () => tryCatch(onFileInputChanged);
    document.getElementById("importButton").onclick = () => tryCatch(importCSV);
  }
});

let globalErrors = new ErrorsHandler();

function onFileInputChanged() {
  const fileInput = document.getElementById("fileInput") as HTMLInputElement;
  const importButton = document.getElementById("importButton") as HTMLButtonElement;

  if (fileInput != null && fileInput.files.length > 0) {
    const file = fileInput.files[0];
    if (file.type === "text/csv") {
      importButton.disabled = false;
      return;
    }
  }

  importButton.disabled = true;
}

async function importCSV() {
  resetAndHideErrors();

  const fileInput = document.getElementById("fileInput") as HTMLInputElement;

  if (fileInput !== null && fileInput.files.length > 0) {
    const file = fileInput.files[0];
    if (file.type !== "text/csv") return;

    Papa.parse<UEDialogueDataTable>(file, {
      header: true,
      skipEmptyLines: true,
      complete: function (result) {
        if (result.errors.length > 0) {
          globalErrors.addParseError(result.errors);
          displayErrorsIfExists();

          result.errors.forEach((error) => {
            console.error("Parsing error: " + error.message);
          });
        }

        Office.context.ui.messageParent(JSON.stringify(result));
      },
    });
  }
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
