# XLSX-Parser

This 'library', if you can even call it that, can be run fully on the browser.
Simply pass it a file and then let it do its "magic". An example can be seen in `App.vue`,

## Usage

```ts
const spreadsheet = await Spreadsheet.loadFromFile(target.files[0]);

for (const worksheet of spreadsheet.worksheets) {
  console.log(`Worksheet ${worksheet.name} has been loaded.`);
  for (const row of worksheet.rows) {
    // You can do something row-related here
    for (const cell of row.cells) {
      // You can do something cell-related here

      // Prints all the cell's values
      console.log(`${cell.value}`);
    }
  }
}
```

It is a bit verbose at the moment, having to jump through so many hoops to get to the **MEAT AND POTATOES** of the spreadsheet, which are the cells. Which is why I plan to (hopefully) create some kind of way to directly get the value of a cell by cell reference, like `B5` as an example, or some other more intuitive way, but I am still thinking about it.

## Disclaimer

This is barely completed and while it has the all the features you'd need (probably) its
still far from complete and is very bare bones (and buggy). If you want to contribute, feel free, I can
already tell this project is one of those where I have a burst of motivation for 6 hours
and dump all my time into this only to forget it the next day.
