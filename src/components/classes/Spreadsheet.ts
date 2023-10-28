import JSZip from 'jszip';
import Converters from '../utils/Converters';
import Workbook from './Workbook';

export default class Spreadsheet {
  workbook!: Workbook;
  sharedStrings: string[] = [];

  public static async loadFromFile(file: File) {
    const zip = await JSZip.loadAsync(file);
    const relevantEntries: {
      workbook: JSZip.JSZipObject | undefined;
      sharedStrings: JSZip.JSZipObject | undefined;
      worksheets: JSZip.JSZipObject[];
    } = {
      workbook: undefined,
      sharedStrings: undefined,
      worksheets: [],
    };

    zip.forEach((_, entry) => {
      if (entry.name.endsWith('workbook.xml')) {
        relevantEntries.workbook = entry;
      }
      if (entry.name.endsWith('sharedStrings.xml')) {
        relevantEntries.sharedStrings = entry;
      }
      if (entry.name.includes('worksheets') && !entry.dir) {
        relevantEntries.worksheets.push(entry);
      }
    });

    const spreadsheet = new Spreadsheet();

    const sharedStringsDocument = await Converters.jsZipObjectToDocument(relevantEntries.sharedStrings!);
    spreadsheet.sharedStrings = Array.from(sharedStringsDocument.querySelectorAll('t')).map((el) => el.innerHTML);

    const workbookDocument = await Converters.jsZipObjectToDocument(relevantEntries.workbook!);
    spreadsheet.workbook = await Workbook.loadAsync(spreadsheet, workbookDocument, relevantEntries.worksheets);

    return spreadsheet;
  }
}
