import JSZip from 'jszip';

export default class Converters {
  static parser: DOMParser;

  static jsZipObjectToDocument = async (jsZipObject: JSZip.JSZipObject): Promise<Document> => {
    if (this.parser == null) {
      this.parser = new DOMParser();
    }
    const fileContent = await jsZipObject.async('text');
    return this.parser.parseFromString(fileContent, 'text/xml');
  };
}
