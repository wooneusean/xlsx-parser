export class CollectionHelper {
  static getFirstOrDefault(collection: HTMLCollection, fallback = null) {
    if (collection == null || collection.length == 0) {
      return fallback;
    }

    return collection[0];
  }
}
