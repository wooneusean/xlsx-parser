import JSZip from 'jszip';
import { Converters } from '../utils/Converters';
import { Spreadsheet } from './Spreadsheet';
import { Worksheet } from './Worksheet';

/**
 * {@link Workbook} class contains the {@link Worksheet}s that are within an Excel file.
 */
export class Workbook {
  worksheets: Worksheet[] = [];
  spreadsheet!: Spreadsheet;

  private constructor() {}

  /**
   * This function is primarily used by the {@link Spreadsheet} class to load from an Excel file.
   * You really don't need to run this function manually.
   * @param spreadsheet {@link Spreadsheet} instance which can be used to traverse up the hiearchy if needed.
   * @param document The {@link Document} of the Workbook, parsed from XML.
   * @param worksheetEntries Worksheet files within the Excel file. Fun fact, an Excel file is just a zip file.
   * @returns {Promise<Workbook>}
   */
  static async loadAsync(
    spreadsheet: Spreadsheet,
    document: Document,
    worksheetEntries: JSZip.JSZipObject[]
  ): Promise<Workbook> {
    const workbook = new Workbook();
    workbook.spreadsheet = spreadsheet;

    const sheets = Array.from(document.querySelectorAll('sheet'));

    for (const sheet of sheets) {
      const sheetRId = sheet.getAttribute('r:id');
      if (sheetRId == null) throw new Error('This sheet does not have an ID.');

      const sheetName = sheet.getAttribute('name') ?? '';

      const sheetId = parseInt(sheetRId.slice(3));

      const worksheetEntry = worksheetEntries[sheetId - 1];

      const worksheetDocument = await Converters.jsZipObjectToDocument(worksheetEntry);

      workbook.worksheets.push(new Worksheet(workbook, sheetName, sheetId, worksheetDocument));
    }

    return workbook;
  }
}
