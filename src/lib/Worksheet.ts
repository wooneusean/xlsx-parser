import { CollectionHelper } from '../utils/CollectionHelper';
import { Constants } from '../utils/Constants';
import { Converters } from '../utils/Converters';
import { Workbook } from './Workbook';

export type TReference = {
  row: number;
  col: number;
};

/**
 *
 */
export class Worksheet {
  name: string;
  sheetId: number;
  isActive: boolean = false;
  activeCellReference: string = '';
  dimensions: { from: TReference; to: TReference } = {} as any;
  cells: { [col: number]: { [row: number]: any } } = {};
  workbook: Workbook;

  constructor(workbook: Workbook, name: string, sheetId: number, document: Document) {
    this.workbook = workbook;
    this.name = name;
    this.sheetId = sheetId;

    this.processDocument(document);
  }

  at(ref: string) {
    const cellRef = Converters.toCellReference(ref);
    return this.cells[cellRef.col][cellRef.row];
  }

  /**
   *
   * @param document
   */
  private processDocument(document: Document) {
    this.isActive =
      CollectionHelper.getFirstOrDefault(document.getElementsByTagName('sheetView'))?.getAttribute('tabSelected') !=
      null;

    this.activeCellReference =
      CollectionHelper.getFirstOrDefault(document.getElementsByTagName('selection'))?.getAttribute('activeCell') ?? '';

    const dims = CollectionHelper.getFirstOrDefault(document.getElementsByTagName('dimension'))?.getAttribute('ref');

    if (dims) {
      const [from, to] = dims.split(':');
      if (from && to) {
        this.dimensions.from = Converters.toCellReference(from);
        this.dimensions.to = Converters.toCellReference(to);
      } else {
        this.dimensions.from = Converters.toCellReference(dims);
        this.dimensions.to = Converters.toCellReference(dims);
      }
    }

    for (let i = 0; i <= this.dimensions.to.col; i++) {
      this.cells[i] = {};
      for (let j = 0; j <= this.dimensions.to.row; j++) {
        this.cells[i][j] = '';
      }
    }

    const rowElements = document.getElementsByTagName('row');

    for (const rowElement of rowElements) {
      const cellElements = rowElement.getElementsByTagName('c');

      for (const cellElement of cellElements) {
        const valueElement = CollectionHelper.getFirstOrDefault(cellElement.getElementsByTagName('v'));

        if (valueElement == null) {
          continue;
        }

        const cellReferenceAttribute = cellElement.getAttribute('r');

        if (cellReferenceAttribute == null) {
          continue;
        }

        const cellReference = Converters.toCellReference(cellReferenceAttribute);
        const cellType = cellElement.getAttribute('t');
        const cellSerial = parseInt(cellElement.getAttribute('s') as string) ?? null;

        let value = '';
        if (cellSerial == 6 && cellType !== 's') {
          let date = new Date(Constants.Epoch);
          date.setDate(date.getDate() + parseInt(valueElement.innerHTML));
          value = date.toISOString();
        } else if (cellType === 's') {
          const sharedStringsIndex = parseInt(valueElement.innerHTML);
          value = this.workbook.spreadsheet.sharedStrings[sharedStringsIndex];
        } else {
          value = valueElement.innerHTML;
        }

        this.cells[cellReference.col][cellReference.row] = value;
      }
    }
  }
}
