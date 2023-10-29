import CollectionHelper from '../utils/CollectionHelper';
import Constants from '../utils/Constants';
import Workbook from './Workbook';

export class WorksheetCell {
  _row: WorksheetRow;
  reference: string;
  type: string | null;
  serial: number | null;
  value: string;

  constructor(element: Element, row: WorksheetRow) {
    this._row = row;
    this.reference = element.getAttribute('r') ?? '';
    this.type = element.getAttribute('t');
    this.serial = parseInt(element.getAttribute('s') as string) ?? null;

    const valueElement = CollectionHelper.getFirstOrDefault(element.getElementsByTagName('v'));
    if (valueElement == null) {
      this.value = '';
      return;
    }

    if (this.serial == 6 && this.type !== 's') {
      let date = new Date(Constants.Epoch);
      date.setDate(date.getDate() + parseInt(valueElement.innerHTML));

      this.value = date.toISOString();
      return;
    }

    if (this.type === 's') {
      const sharedStringsIndex = parseInt(valueElement.innerHTML);
      this.value = this._row._worksheet._workbook._spreadsheet.sharedStrings[sharedStringsIndex];
    } else {
      this.value = valueElement.innerHTML;
    }
  }
}

export class WorksheetRow {
  _worksheet: Worksheet;
  index: number;
  cells: WorksheetCell[] = [];
  spans: string;

  constructor(element: Element, worksheet: Worksheet) {
    this._worksheet = worksheet;
    this.index = parseInt(element.getAttribute('r') ?? '0');
    this.spans = element.getAttribute('spans') ?? '';
    const cellElements = element.getElementsByTagName('c');
    for (const cellElement of cellElements) {
      this.cells.push(new WorksheetCell(cellElement, this));
    }
  }
}

export default class Worksheet {
  _workbook: Workbook;
  name: string;
  sheetId: number;
  isActive: boolean = false;
  activeCellReference: string = '';
  dimensions: string = '';
  rows: WorksheetRow[] = [];
  cells: Map<string, WorksheetCell> = new Map();

  constructor(workbook: Workbook, name: string, sheetId: number, document: Document) {
    this._workbook = workbook;
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
      const row = new WorksheetRow(rowElement, this);
      const rowCells = row.cells;
      this.rows.push(row);
      for (const cell of rowCells) {
        this.cells.set(cell.reference, cell);
      }
    }
  }

  public static getCellReference(reference: string): { columnRef: string; rowRef: number } {
    const referenceMatch = /^(\w+)(\d+)$/.exec(reference);
    if (referenceMatch == null) {
      throw new Error('This is not a cell reference');
    }

    const columnRef = referenceMatch.groups![1];
    const rowRef = parseInt(referenceMatch.groups![2]);
    return { columnRef, rowRef };
  }

  at(reference: string) {
    return this.cells.get(reference);
  }
}
