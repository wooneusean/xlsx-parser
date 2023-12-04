import JSZip from 'jszip';
import { TReference } from '../lib/Worksheet';

export class Converters {
  static parser: DOMParser;

  static jsZipObjectToDocument = async (jsZipObject: JSZip.JSZipObject): Promise<Document> => {
    if (this.parser == null) {
      this.parser = new DOMParser();
    }
    const fileContent = await jsZipObject.async('text');
    return this.parser.parseFromString(fileContent, 'text/xml');
  };

  static excelColumnToNumber = (col: string) => {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result *= 26; // Multiply the result by 26 for each character processed
      result += col.charCodeAt(i) - 'A'.charCodeAt(0) + 1; // Add the value of the current character
    }
    return result - 1; // Subtract 1 to convert to zero-based index
  };

  static toCellReference = (refString: string): TReference => {
    const ref = /([a-zA-Z]+)([0-9]+)/.exec(refString);

    if (ref == null) return { col: -1, row: -1 };

    return { col: this.excelColumnToNumber(ref[1]), row: parseInt(ref[2]) - 1 };
  };
}
