import Papa from "papaparse";

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

export interface ErrorsHandler {
  addinErrors: string[];
  parseErrors: Papa.ParseError[];
}
