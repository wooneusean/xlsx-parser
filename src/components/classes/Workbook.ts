import JSZip from 'jszip';
import Converters from '../utils/Converters';
import Spreadsheet from './Spreadsheet';
import Worksheet from './Worksheet';

export default class Workbook {
  __spreadsheet: Spreadsheet;
  worksheets: Worksheet[] = [];

  constructor(spreadsheet: Spreadsheet) {
    this.__spreadsheet = spreadsheet;
  }

  static async loadAsync(spreadsheet: Spreadsheet, document: Document, worksheetEntries: JSZip.JSZipObject[]) {
    const workbook = new Workbook(spreadsheet);

    const sheets = Array.from(document.querySelectorAll('sheet'));

    for (const sheet of sheets) {
      const sheetId = parseInt(sheet.getAttribute('sheetId') ?? '0');
      const sheetName = sheet.getAttribute('name') ?? '';

      const worksheetEntry = worksheetEntries[sheetId - 1];

      const worksheetDocument = await Converters.jsZipObjectToDocument(worksheetEntry);

      workbook.worksheets.push(new Worksheet(workbook, sheetName, sheetId, worksheetDocument));
    }

    return workbook;
  }
}
