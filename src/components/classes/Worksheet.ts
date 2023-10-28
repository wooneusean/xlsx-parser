import CollectionHelper from '../utils/CollectionHelper';
import Workbook from './Workbook';

export type WorksheetCellType = 'n' | 's';

export class WorksheetCell {
  __row: WorksheetRow;
  reference: string;
  type: WorksheetCellType;
  value: string;

  constructor(element: Element, row: WorksheetRow) {
    this.__row = row;
    this.reference = element.getAttribute('r') ?? '';
    this.type = (element.getAttribute('t') ?? 'n') as WorksheetCellType;
    if (this.type === 's') {
      const sharedStringsIndex = parseInt(element.getElementsByTagName('v')[0].innerHTML);
      this.value = this.__row.__worksheet.__workbook.__spreadsheet.sharedStrings[sharedStringsIndex];
    } else {
      this.value = element.getElementsByTagName('v')[0].innerHTML;
    }
  }
}

export class WorksheetRow {
  __worksheet: Worksheet;
  index: number;
  cells: WorksheetCell[] = [];
  spans: string;

  constructor(element: Element, worksheet: Worksheet) {
    this.__worksheet = worksheet;
    this.index = parseInt(element.getAttribute('r') ?? '0');
    this.spans = element.getAttribute('spans') ?? '';
    const cellElements = element.getElementsByTagName('c');
    for (const cellElement of cellElements) {
      this.cells.push(new WorksheetCell(cellElement, this));
    }
  }
}

export default class Worksheet {
  __workbook: Workbook;
  name: string;
  sheetId: number;
  isActive: boolean = false;
  activeCellReference: string = '';
  dimensions: string = '';
  rows: WorksheetRow[] = [];

  constructor(workbook: Workbook, name: string, sheetId: number, document: Document) {
    this.__workbook = workbook;
    this.name = name;
    this.sheetId = sheetId;

    this.processDocument(document);
  }

  private processDocument(document: Document) {
    this.isActive =
      CollectionHelper.getFirstOrDefault(document.getElementsByTagName('sheetView'))?.getAttribute('tabSelected') !=
      null;

    this.activeCellReference =
      CollectionHelper.getFirstOrDefault(document.getElementsByTagName('selection'))?.getAttribute('activeCell') ?? '';

    this.dimensions =
      CollectionHelper.getFirstOrDefault(document.getElementsByTagName('dimension'))?.getAttribute('ref') ?? '';

    const rowElements = document.getElementsByTagName('row');
    for (const rowElement of rowElements) {
      this.rows.push(new WorksheetRow(rowElement, this));
    }
  }
}
