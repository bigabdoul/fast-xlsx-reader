const xlsx = require("xlsx");
const fs = require("fs");
const {
  tryConvertDate
} = require("./xldates");

/**
 * Reads Excel (.xlsx) sheets in a more efficient way.
 */
class FastXlsxReader {
  /**
   * Initialize a new instance of the FastXlsxReader class.
   * @param {object} options An object with processing options.
   * Supported interface is:
   * {
   * input: string;
   * output?: string|WriteStream|null;
   * format?: string;
   * sheetname?: string;
   * hasHeader?: boolean;
   * headerPrefix?: string;
   * lowerCaseHeaders?: boolean;
   * schema?: object;
   * onHeader?: header => void;
   * onCell?: (cell, rowIndex, colIndex) => void;
   * onRecord?: (record, index) => void;
   * onFinish?: (items, rowsProcessed) => void;
   * onError?: err => void;
   * useMemoryForItems?: boolean
   * }
   */
  constructor(options) {
    this.options = options;
    this._header = null;
    this._hasSchemaErr = false;
    this._outStream = null;
    this._hasRecord = false;
    this._items = [];
    this._eventHandlers = {};
    this._eventHandlerCount = 0;
  }

  /** Does a valid output stream exist? */
  get hasStream() {
    return this._outStream !== null && this._outStream !== undefined;
  }

  /** Is the output format "json"? */
  get isJson() {
    return this.options.format === "json";
  }

  /** Get the header. */
  get header() {
    return this._header;
  }

  /** Get the current row object. */
  get currentRow() {
    return this._currentRow;
  }

  /**
   * Add an event handler.
   * @param {string} event The name of the event to listen to.
   * @param {Function} handler A callback function to handle the event.
   */
  on(event, handler) {
    const {
      _eventHandlers
    } = this;

    if (!(event in _eventHandlers)) {
      _eventHandlers[event] = [];
    }

    _eventHandlers[event].push(handler);
    this._eventHandlerCount++;
  }

  /**
   * Remove an event handler, or all handlers for the specified event.
   * @param {string} event The name of the event to remove.
   * @param {Function} handler Optional: The callback function to remove.
   */
  off(event, handler) {
    const {
      _eventHandlers
    } = this;

    if (event in _eventHandlers) {
      const e_array = _eventHandlers[event];
      if (e_array instanceof Array) {
        if (handler) {
          const idx = e_array.indexOf(handler);

          if (idx > -1) {
            e_array.splice(idx, 1);
            this._eventHandlerCount--;
          }
        } else {
          delete _eventHandlers[event];
          this._eventHandlerCount -= e_array.length;
        }
      }
    }
  }

  /**
   * Read sequentially the rows contained in the
   * Excel sheet (specified in the options).
   */
  read() {
    const {
      input,
      sheetname,
      schema,
      lowerCaseHeaders,
      onRecord,
      onCell,
      onError,
      useMemoryForItems // useful when no onRecord handler and no output provided
    } = this.options;

    console.log("Reading Excel file...", input);

    const eventNames = [...DEFAULT_EVENTS];
    if (typeof onCell === "function") eventNames.push("cell");

    this._rowsProcessed = 0;
    this._hasRecord = false;
    this._createOutStream();
    this._writeHeader();

    FastXlsxReader.iterate(
      input,
      sheetname,
      (eventName, data, rowIndex, colIndex) => {
        switch (eventName) {
          case "start":
            break;
          case "cell":
            onCell.call(this, data, rowIndex, colIndex);
            break;
          case "record":
            if (rowIndex === 0) {
              this._readHeader(data);
              data = this._header;
            } else {
              data = this._readRow(
                data,
                rowIndex,
                schema,
                lowerCaseHeaders,
                onRecord,
                onError,
                useMemoryForItems
              );
            }
            break;
          case "end":
            this._finalize();
            break;
          case "error":
            this._quit(1, data);
            break;
          default:
            break;
        }

        if (this._eventHandlerCount > 0)
          this._fireEvent(eventName, data, rowIndex, colIndex);
      },
      eventNames,
      this // thisArg for the callback
    );
  }

  _fireEvent(name, data, rowIndex, colIndex) {
    const callbacks = this._eventHandlers[name];
    if (callbacks instanceof Array) {
      for (let i = 0; i < callbacks.length; i++) {
        callbacks[i].call(this, data, rowIndex, colIndex);
      }
    }
  }

  _readHeader(row) {
    const {
      schema,
      lowerCaseHeaders: lowerCase,
      hasHeader = true,
      headerPrefix = "header_",
      onHeader
    } = this.options;

    if (!hasHeader) {
      this._header = [];
      if (!!schema) {
        // create header from the provided schema
        for (const key in schema) {
          this._header.push(key);
        }
      } else {
        // no schema, give arbitrary header names
        for (let i = 1; i <= row.length; i++) {
          const name = headerPrefix + i;
          this._header.push(lowerCase ? name.toLowerCase() : name);
        }
        console.log("Created arbitrary header: ", this._header);
      }
    } else {
      this._header = FastXlsxReader._normalizeHeader(row, !schema && lowerCase);
    }

    this._rowsProcessed++; // the header is a row, so it counts

    if (typeof onHeader === "function") {
      onHeader.call(this, this._header);
    }
  }

  _readRow(
    row,
    index,
    schema,
    lowerCase,
    onRecord,
    onError,
    useMemoryForItems
  ) {
    let record;

    if (!!schema) {
      record = this._rowFromSchema(row, schema, onError);
    } else {
      record = this._rowFromHeader(row, lowerCase);
    }

    this._rowsProcessed++;
    this._currentRow = record;

    if (typeof onRecord === "function") onRecord.call(this, record, index);
    else if (useMemoryForItems && !this.hasStream) this._items.push(record);

    this._writeRecord(record);
    return record;
  }

  _rowFromSchema(row, schema, onError) {
    const obj = {};
    this._header.forEach((column, index) => {
      const meta = schema[column];
      if (meta) {
        const {
          prop,
          type: cast
        } = meta;
        let value = row[index];
        if (value === undefined || value === null)
          value = "";
        if (typeof cast === "function") {
          if (cast.prototype.constructor.name === "Date") {
            obj[prop] = tryConvertDate(value);
          } else
            obj[prop] = cast(value);
        } else obj[prop] = value;
      } else if (!this._hasSchemaErr) {
        this._hasSchemaErr = true;
        const msg = `#ERR_SCHEMA: Invalid schema! No mapping for column "${column}".`;
        onError && onError.call(this, msg);
        console.error(msg, schema);
        this._quit();
      }
    });
    return obj;
  }

  _rowFromHeader(row, lowerCase) {
    const obj = {};
    this._header.forEach((column, index) => {
      const key = lowerCase ? column.toLowerCase() : column;
      obj[key] = row[index];
    });
    return obj;
  }

  /**
   * Create a WriteStream instance using the 'output' property of options.
   * The output property can be a string or an instance of WriteStream.
   * It can also be undefined or null.
   */
  _createOutStream() {
    if (!this.hasStream) {
      const {
        output
      } = this.options;

      if (output === undefined || output === null) return;

      if (typeof output === "string")
        this._outStream = fs.createWriteStream(output, {
          flags: "w"
        });
      else if (typeof output.write === "function")
        // we assume output is a WriteStream instance
        this._outStream = output;
      else
        throw new Error(
          "Output must be a string or an instance of WriteStream."
        );
    }
  }

  _writeHeader() {
    if (this.hasStream) {
      if (this.isJson) this._outStream.write("[");
    }
  }

  _writeRecord(record) {
    if (this.hasStream) {
      if (this.isJson) {
        if (this._hasRecord) this._outStream.write(",");
        else this._hasRecord = true;

        this._outStream.write(JSON.stringify(record));
      }
    }
  }

  _writeFooter() {
    if (this.hasStream) {
      if (this.isJson) {
        this._outStream.end("]");
      }
      this._outStream.close();
      this._outStream = null;
    }
  }

  _finalize() {
    this._writeFooter();

    const {
      onFinish,
      onRecord
    } = this.options;

    if (typeof onFinish === "function") {
      if (typeof onRecord !== "function" && !this.hasStream) {
        // when writing to a file, use the 'close' event
        // the 'end' event may fire before the file has been written
        onFinish.call(this, this._items, this._rowsProcessed);
      } else {
        onFinish.call(this, null, this._rowsProcessed);
      }
    }
  }

  _quit(code, error) {
    this._writeFooter();
    const {
      onError
    } = this.options;
    if (typeof onError === "function") onError.call(this, error);
    else console.error(`#ERR: ${error}`);
    process.exit(code || 1);
  }

  static _normalizeHeader(row, lowerCase) {
    row.forEach((column, index) => {
      let col = column.toString().trim();
      if (lowerCase) col = col.toLowerCase();
      row[index] = col;
    });
    return row;
  }

  /**
   * Iterate over all rows contained in an Excel sheet.
   * @param {string} input The Excel input file name.
   * @param {string} sheetname Optional: The name of the sheet to iterate over.
   * If undefined, use the first sheet.
   * @param {Function} callback A  function to invoke based on different events.
   * @param {string|Array<string>} eventNames Optional: A one-dimensional array or a
   * comma-separated list of event names to invoke on callback.
   * If undefined, default events ("start", "record", "end", "error") are used.
   * @param {any} thisArg The object to be used as the current object when invoking
   * the callback.
   */
  static iterate(input, sheetname, callback, eventNames, thisArg) {
    let events = eventNames;
    try {
      const book = xlsx.readFile(input);
      const sheet = sheetname ?
        book.Sheets[sheetname] :
        book.Sheets[book.SheetNames[0]];

      const range = xlsx.utils.decode_range(sheet["!ref"]);
      const {
        r: startRow,
        c: startCol
      } = range.s;
      const {
        r: endRow,
        c: endCol
      } = range.e;

      if (eventNames instanceof String) events = eventNames.split(",");
      else if (eventNames instanceof Array)
        events = eventNames.filter(e => e instanceof String);
      else events = [...DEFAULT_EVENTS];

      const noEvents = events.length === 0;
      const hasStartCb = noEvents || events.indexOf("start") > -1;
      const hasRecordCb = noEvents || events.indexOf("record") > -1;
      const hasEndCb = noEvents || events.indexOf("end") > -1;
      const hasCellCb = events.indexOf("cell") > -1; // must be explicitly set

      let rowCount = 0;

      if (hasStartCb) callback.call(thisArg, "start");

      for (let row_index = startRow; row_index <= endRow; row_index++) {
        const row = [];
        for (let col_index = startCol; col_index <= endCol; col_index++) {
          const encodedCell = xlsx.utils.encode_cell({
            r: row_index,
            c: col_index
          });
          const cell = sheet[encodedCell];
          if (!!cell) {
            /* Sample cell contents:
            (number) { t: 'n', v: 269, w: '269' }
            (number) { t: 'n', v: 421, w: '421' }
            (string) { t: 's', v: 'hello', r: '<t>hello</t>', h: 'hello', w: 'hello' }
            (string) { t: 's', v: 'world!', r: '<t>world!</t>', h: 'world!', w: 'world!' }
            */
            row.push(cell.v);
            if (hasCellCb)
              callback.call(thisArg, "cell", cell, row_index, col_index);
          } else {
            row.push(cell);
          }
        }
        rowCount++;
        if (hasRecordCb) callback.call(thisArg, "record", row, row_index);
      }

      if (hasEndCb) callback.call(thisArg, "end", rowCount);
    } catch (error) {
      if (events.indexOf("error") > -1) callback.call(thisArg, "error", error);
      else {
        if (typeof error === "string") throw new Error(error);
        else throw error;
      }
    }
  }
}

const DEFAULT_EVENTS = ["start", "record", "end", "error"];

module.exports = FastXlsxReader;