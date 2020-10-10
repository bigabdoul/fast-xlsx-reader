# fast-xlsx-reader

`fast-xlsx-reader` is a node package (https://npmjs.com/package/fast-xlsx-reader) based on the famous and awesome `xlsx`
(https://npmjs.com/package/xlsx) package.

It's used for reading row-by-row an Excel worksheet, and optionally writing
sequentially each row to a stream in JSON format, or whatever format you desire.

It was written to overcome some limitations of similar packages such as:

- `read-excel-file` (see https://github.com/catamphetamine/read-excel-file/issues/38) and
- `xlsx-to-json-lc` (https://npmjs.com/package/xlsx-to-json-lc): limited to bulk processing
  an entire worksheet.

## Install

Installation is done using the npm install command:

```
$ npm i fast-xlsx-reader
```

## Usage

### New: Introducing `FastXlsxSheetReader`

You can now create an instance of the new `FastXlsxSheetReader` class addition 
and read rows at your own pace. This is an easier alternative of using the library
with new capabilities. Here are some code snippets that show how it's done:

```JavaScript
const excel = require("fast-xlsx-reader");

// Create an instance of the FastXlsxSheetReader class
const reader = excel.createReader({
    input: input_file
});

// log some metadata
console.log('\nSheet properties:');
console.log('-------------------')
console.log(`Row start: \t${reader.startRow}`);
console.log(`Column start: \t${reader.startCol}`);
console.log(`Row end: \t${reader.endRow}`);
console.log(`Column end: \t${reader.endCol}`);
console.log(`Number of rows: ${reader.rowCount}`);
console.log('');

let row = reader.readNext();
console.log(`Row #${reader.rowIndex + 1}: `, row);

row = reader.readNext();
console.log(`Row #${reader.rowIndex + 1}: `, row);

// There're 2 ways to iterate over the rows:
// 1st: readNext() returns a row
while ((row = reader.readNext())) {
    console.log(`Row #${reader.rowIndex + 1}: `, row);
}

console.log('\nReset the reader');
reader.reset();

// 2nd: moveNext() returns true or false
while (reader.moveNext()) {
    // use FastXlsxSheetReader.current property for the current row
    console.log(`Row #${reader.rowIndex + 1}:`, reader.current);
}

console.log("\nReset and read next...");
row = reader.reset().readNext();
console.log(`Row #${reader.rowIndex + 1}: `, row);

const startIndex = 3;
let count = 4;
const rows = reader.readMany(startIndex, count);
console.log(`Read ${count} rows at once:`, rows);
console.log('Read next:', reader.readNext());

// with event callbacks
reader.reset().oncell = (cell, col, row) => {
    console.log(`Cell[${col}, ${row}] =`, cell);
};

// this event handler is required when calling readAll()
reader.on("record", (data, index) => {
    console.log(`\nThe above was row #${index + 1}:`, data);
    console.log('');
});

// read 3 rows backwards starting at one row above the last (index -2)
// (the last row is at index -1)
count = 3;
console.log(`Read ${count} rows backwards:`, reader.readMany(-2, count));

// read cell
let colidx = 3,
    rowidx = 2;
console.log(`Reading cell[${colidx}, ${rowidx}]:`, reader.readCell(colidx, rowidx));

// To read all rows, make sure that the 'record' event handler is defined!
console.log("Reading all rows...");

// event raised before reading each row
reader.on("beforerecord", (index) => {
    console.log(`Reading row #${index + 1}`);
});

// pass true if you want to read backwards;
// read all rows by using a 'record' callback (the previously registered
// 'record' event handler is ignored)
const rowCount = reader.readAll( /* backwards: */ false, (row, index) => {
    index.toString();
    console.log("Data", row);
    console.log('');
});

console.log(`Done! ${rowCount} rows read.`);

// clean up
reader.destroy();

// reader.reset().readNext(); // throws an error because we destroyed the reader
```

### Legacy usage with written output (file or WriteStream)

```JavaScript
const path = require("path");
const excel = require("fast-xlsx-reader");
const schema_file = require("./data/your-schema-for-excel-file-to-read.js");

// assuming that all files are located under ./data
const datadir = path.join(__dirname, "./data");
const timeStart = Date.now(); // just to have an idea about the processing time
const input_file = path.join(datadir, "excel-file-to-read.xlsx");
const output_file = path.join(datadir, "file-to-write-to.json");

// specify read options
const options = {
    input: input_file,
    output: output_file,  // can be a NodeJs WriteStream
    format: "json", // default and only built-in supported format for now
    sheetname: "sheet_to_read", // specify the name of the sheet you wanna read
    schema: schema_file, // see below for a sample schema file
    hasHeader: true,  // we think that the first row has a header (default)
    // useMemoryForItems: true,

    // this callback is not required when writing the output in json format
    onRecord: (record, index) => {
      // act on each row; use this callback to do something useful
      // record: a JSON object
      // index: the zero-based row number
    },

    onFinish: (items, rowsProcessed) => {
        // the 'items' parameter represents the in-memory JSON array;
        // it's an empty array when 'onRecord' and 'output' are not set;
        // to use it despite these non-provided options, uncomment
        // the line: useMemoryForItems: true,
        console.log(`Processed ${rowsProcessed} rows in ${Date.now() - timeStart}ms`);
        console.log(`Output written to ${output_file}`);
    }
}

// read all rows in the worksheet
excel.read(options);
```

## Using a schema

Sample schema: `your-schema-for-excel-file-to-read.js` (whatever)

```JavaScript
module.exports = {
    "REGION_CODE": {
        prop: "rec"
    },
    "REGION": {
        prop: "re"
    },
    "CIRCLE_CODE": {
        prop: "cic"
    },
    "CIRCLE": {
        prop: "ci"
    },
    "COMMUNE_CODE": {
        prop: "coc"
    },
    "COMMUNE": {
        prop: "co"
    },
    "CENTER_CODE": {
        prop: "cec"
    },
    "CENTER": {
        prop: "ce"
    },
    "BUREAU_CODE": {
        prop: "buc"
    },
    "BUREAU": {
        prop: "bu"
    },
    "MEN": {
        prop: "m"
    },
    "WOMEN": {
        prop: "w"
    },
    "TOTAL_VOTERS": {
        prop: "vo"
    }
}
```

The above schema maps the Excel WorkSheet's columns `REGION_CODE`, `REGION`, etc. to their respective JSON properties `rec`, `re`, etc. This implies that each row will be presented as a JSON object like so:

```JavaScript
{
  rec: "01",
  re: "Region 01",
  cic: "01",
  ci: "Circle 01",
  coc: "01",
  co: "Commune 01",
  cec: "01",
  ce: "Center 01",
  buc: "001-001",
  bu: "Voting Bureau 001-001",
  m: 135,
  w: 110,
  vo: 245
}
```

You can optionally specify the type of each column if you want to do custom conversions. This is not required because the types exported exactly match those in the Worksheet's cells. For instance, you can use a schema like this:

```JavaScript
{
    "REGION_CODE": {
        prop: "rec",
        type: Number
    },
    "REGION": {
        prop: "re"
    },
    "CIRCLE_CODE": {
        prop: "cic",
        type: Number
    },
    "CIRCLE": {
        prop: "ci"
    },
    "COMMUNE_CODE": {
        prop: "coc",
        type: Number
    },
    "DATE_CREATED": {
        prop: "dat",
        type: Date
    },
    "CENTER_CODE": {
        prop: "cec",
        type: String
    },
    "CENTER": {
        prop: "ce",
        type: String
    }
}
```

## Features

- Row-by-row reading.
- Efficient memory usage.
- Use a schema to transform header columns.
- Use callbacks or events to act on header columns, each cell, or every single row.

## Quick Start

Create the project folder structure:

```
$ mkdir my-excel-app && cd my-excel-app
$ mkdir src && cd src
```

Install dependencies:

```
$ npm init --yes
$ npm install fast-xlsx-reader
```

Under the `src` folder, create a file named `main.js` (or whatever), and paste the below usage code in it:

```JavaScript
const path = require("path");
const excel = require("fast-xlsx-reader");
const schema_file = require("./data/your-schema-for-excel-file-to-read.js");

const timeStart = Date.now();
const datadir = path.join(__dirname, "./data");
const input_file = path.join(datadir, "excel-file-to-read.xlsx");
const output_file = path.join(datadir, "file-to-write-to.json");

const options = {
    input: input_file,
    output: output_file,
    format: "json",
    sheetname: "sheet_to_read",
    schema: schema_file,

    // adding a 'cell' event handler can significantly slow down the reader
    onCell: (cell, rowIndex, colIndex) => {
        // cell is an object of the xlsx module.

        /* sample cell contents:
        (number) { t: 'n', v: 269, w: '269' }
        (number) { t: 'n', v: 421, w: '421' }
        (string) { t: 's', v: 'hello', r: '<t>hello</t>', h: 'hello', w: 'hello' }
        (string) { t: 's', v: 'world!', r: '<t>world!</t>', h: 'world!', w: 'world!' }
        */

        // the 'v' property is strongly-typed
        if (cell.v === 269) {
            // do something with the cell;
        }

        // the 'w' property is the string representation
        if (cell.w === "269") {
            // will work
        }

        if (cell.w === 269) {
            // will not work
        }
    },

    onFinish: (_, rowsProcessed) => {
        console.log(`Processed ${rowsProcessed} rows in ${Date.now() - timeStart}ms`);
        console.log(`Output written to ${output_file}`);
    }
}

excel.read(options);
```

Directly under the `src` folder, create a folder named `data` and place an Excel file named `excel-file-to-read.xlsx` in it.

Of course, you can name the folder and file whatever you want; just bear in mind to change their names accordingly in the code snippet above.

Also, create a schema file named `your-schema-for-excel-file-to-read.js` (or whatever) in the `data` directory. Using a schema file is not mandatory. If you do, just make sure to use something that matches the columns in the sheet you want to read.

From within the current `src` folder, run:

```
$ node main.js
```

Look in the `data` directory to see the generated file `file-to-write-to.json` (or whatever you called it).
