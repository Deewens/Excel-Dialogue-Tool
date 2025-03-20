interface Config {
  worksheetName: string;
  tableName: string;
  tableHeaderRows: string[];
}

export const config: Config = {
  worksheetName: "Dialogue Editor",
  tableName: "DialogueTable",
  tableHeaderRows: [
    "ID",
    "Speaker",
    "Text",
    "NextLineID",
    "Choices",
    "Conditions",
    "LocNamespace",
    "LocKey",
    "SpeakerData",
  ],
};
