import Papa from "papaparse";

export class ErrorsHandler {
  addinErrors: string[];
  parseErrors: Papa.ParseError[];

  addAddinError(errors: string[]) {
    this.addinErrors.push(...errors);
  }

  addParseError(errors: Papa.ParseError[]) {
    this.parseErrors.push(...errors);
  }

  clear() {
    this.addinErrors = [];
    this.parseErrors = [];
  }
}
