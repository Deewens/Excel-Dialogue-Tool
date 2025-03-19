import Papa from "papaparse";
import { ErrorsHandler } from "./errors-handler";

export interface UEDialogueDataTable {
  "---": string;
  Speaker: string;
  DialogueText: string;
  Conditions: string;
  NextLineID: string;
  Choices: string;
  SpeakerData: string;
}

export interface DialogueTable {
  ID: string;
  Speaker: string;
  Text: string;
  LocNamespace: string;
  LocKey: string;
  Conditions: string;
  NextLineID: string;
  Choices: string;
  SpeakerData: string;
}

export interface ImportDialogMessage<T> {
  parseResult: Papa.ParseResult<T>;
}
