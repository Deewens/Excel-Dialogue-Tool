import Papa from "papaparse";
import { UEDialogueDataTable } from "./types";

const REGEX_FTEXT_EXTRACTION = /"(.*?[^\\])"/g;

export async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

/**
 * Converts a 2D array from an Excel Range values into a generic JSON object.
 * @param values
 */
export function returnObjectFromValues<T>(values: string[][]): T[] {
  let objectArray: T[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i];
      continue;
    }

    let object: { [key: string]: string } = {};
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j];
    }

    objectArray.push(object as unknown as T);
  }

  return objectArray;
}

export function extractFTextComponents(FTextString: string) {
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

export function parseCSV(file: File, complete) {
  Papa.parse<UEDialogueDataTable>(file, {
    header: true,
    skipEmptyLines: true,
    complete: function (results) {
      if (results.errors.length > 0) {
        results.errors.forEach((error) => {
          // TODO: handle these errors in a better way
          console.error("Parsing error: " + error.message);
        });
      }

      complete(results.data, results.meta, results.errors);
    },
  });
}
