import Papa from "papaparse";

export interface UEDialogueDataTable {
  "---": string;
  Speaker: string;
  DialogueText: string;
  NextLineID: string;
  Choices: string;
  PreEvent: string;
  PostEvent: string;
  Conditions: string;
  SpeakerData: string;
}

export interface DialogueTable {
  ID: string;
  Speaker: string;
  Text: string;
  NextLineID: string;
  Choices: string;
  PreEvent: string;
  PostEvent: string;
  Conditions: string;
  LocNamespace: string;
  LocKey: string;
  SpeakerData: string;
}

export interface ImportDialogMessage<T> {
  parseResult: Papa.ParseResult<T>;
}
