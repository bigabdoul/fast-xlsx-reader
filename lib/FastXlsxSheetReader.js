const xlsx = require("xlsx");

/**
 * Represents an object that provides methods for reading an Excel sheet.
 */
class FastXlsxSheetReader {
    /**
     * Initialize a new instance of the FastXlsxSheetReader class.
     * @param {string} filename The fully-qualified name of the file to read.
     * @param {string|number} sheetnameOrIndex The name or index of the worksheet to read.
     * If not specified, the first available worksheet will be used.
     * @param {object} thisArg An object that provides context when calling event handlers.
     */
    constructor(filename, sheetnameOrIndex, thisArg) {

        // read Excel file as workbook
        const book = xlsx.readFile(filename);

        this._sheet = typeof sheetnameOrIndex === "string" ?
            book.Sheets[sheetnameOrIndex] :
            (typeof sheetnameOrIndex === "number" ?
                book.Sheets[sheetnameOrIndex] :
                book.Sheets[book.SheetNames[0]]);

        const range = xlsx.utils.decode_range(this._sheet["!ref"]);

        const {
            r: startRow,
            c: startCol
        } = range.s;
        const {
            r: endRow,
            c: endCol
        } = range.e;

        this._rowIndex = startRow - 1;
        this._startRow = startRow;
        this._startCol = startCol;
        this._endRow = endRow;
        this._endCol = endCol;
        this._thisArg = thisArg || this;
    }

    /** Returns the "start" event handler. */
    get onstart() {
        return this._onstart;
    }

    /** Sets the "start" event handler. */
    set onstart(value) {
        this._onstart = _ensureFunction(value);
    }

    /** Returns the "cell" event handler. */
    get oncell() {
        return this._oncell;
    }

    /** Sets the "cell" event handler. */
    set oncell(value) {
        this._oncell = _ensureFunction(value);
    }

    /** Returns the "beforerecord" event handler. */
    get onbeforerecord() {
        return this._onbeforerecord;
    }

    /** Sets the "beforerecord" event handler. */
    set onbeforerecord(value) {
        this._onbeforerecord = _ensureFunction(value);
    }

    /** Returns the "record" event handler. */
    get onrecord() {
        return this._onrecord;
    }

    /** Sets the "record" event handler. */
    set onrecord(value) {
        this._onrecord = _ensureFunction(value);
    }

    /** Returns the "end" event handler. */
    get onend() {
        return this._onend;
    }

    /** Sets the "end" event handler. */
    set onend(value) {
        this._onend = _ensureFunction(value);
    }

    /** Returns the "error" event handler. */
    get onerror() {
        return this._onerror;
    }

    /** Sets the "error" event handler. */
    set onerror(value) {
        this._onerror = _ensureFunction(value);
    }

    /** Returns the index of the start row. */
    get startRow() {
        return this._startRow;
    }

    /** Returns the index of the end row. */
    get endRow() {
        return this._endRow;
    }

    /** Returns the index of the start column. */
    get startCol() {
        return this._startCol;
    }

    /** Returns the index of the end column. */
    get endCol() {
        return this._endCol;
    }

    /** Returns the current row index. */
    get rowIndex() {
        return this._rowIndex;
    }

    /** Returns the number of rows. */
    get rowCount() {
        return this._endRow - this._startRow + 1;
    }

    /** Returns the currently opened WorkSheet. */
    get sheet() {
        return this._sheet;
    }

    /**
     * Return the current row.
     * @returns {any[]} A one-dimensional array representing the current row read.
     */
    get current() {
        if (this._currentRow === undefined) {
            return this.read(this._rowIndex);
        }
        return this._currentRow;
    }

    /**
     * Set the row index to -1;
     * @returns {FastXlsxSheetReader} A reference to the current FastXlsxSheetReader instance.
     */
    reset() {
        if (!this._isDestroyed())
            this._rowIndex = this._startRow - 1;
        return this;
    }

    /**
     * Attempts to move to the next row, if any.
     * @returns {boolean} true if the next row can be read, otherwise false.
     */
    moveNext() {
        if (this._rowIndex + 1 <= this._endRow) {
            this._currentRow = this.read(++this._rowIndex);
            return true;
        }
        return false;
    }

    /**
     * Attempt to read the next row.
     * @param {Function} onrecord An optional function to call after reading each row.
     * @returns {object|null} An object if a row was read, otherwise, null.
     */
    readNext(onrecord) {
        if (this.moveNext()) {
            return this.current;
        }
        return null;
    }

    /**
     * Read all rows in the current work sheet.
     * @param {boolean} backwards true to start from the last row, otherwise false.
     * @param {Function} onrecord An optional function to call back when a row is read.
     * If specified, this method takes precedence over the current 'record' event handler.
     * If not specified then a 'record' event handler must exist.
     * @returns {number} A number that represents the number of rows read.
     */
    readAll(backwards, onrecord) {
        onrecord || (onrecord = this.onrecord);

        if (typeof onrecord !== "function") {
            this._handleError(new Error("A callback function for the 'record' event must be specified."));
            return 0;
        }

        let rowCount = 0;
        const thisArg = this._thisArg;

        if (!!this.onstart) this.onstart.call(thisArg);

        if (!!backwards) {
            for (this._rowIndex = this._endRow; this._rowIndex > -1; this._rowIndex--) {
                this.read(this._rowIndex, onrecord);
                rowCount++;
            }
        } else {
            for (this._rowIndex = this._startRow; this._rowIndex <= this._endRow; this._rowIndex++) {
                this.read(this._rowIndex, onrecord);
                rowCount++;
            }
        }

        if (!!this.onend) this.onend.call(thisArg, rowCount);
        return rowCount;
    }

    /**
     * Read the row at the specified index.
     * @param {number} index The relative index of the row to read. Can be negative.
     * @param {Function} onrecord An optional function to call after reading each row.
     * @returns {any[]} A one-dimensional array representing the row read.
     */
    read(index, onrecord) {
        const sheet = this._sheet;

        if (this._isDestroyed())
            return 0;

        index = index | 0;

        if (index < 0)
            index = this._endRow + index;

        if (index < 0 || index > this._endRow) {
            this._handleError(new Error(`The absolute value of the index must be between 0 and ${this._endRow} inclusive.`));
            return null;
        }

        onrecord || (onrecord = this.onrecord);

        const row = [];
        const hasRecordCb = typeof onrecord === "function";
        const oncell = this.oncell;
        const hasCellCb = typeof oncell === "function";
        const thisArg = this._thisArg;

        if (typeof this.onbeforerecord === "function") {
            this.onbeforerecord.call(thisArg, index);
        }

        for (this._colIndex = this._startCol; this._colIndex <= this._endCol; this._colIndex++) {
            const encodedCell = xlsx.utils.encode_cell({
                r: index,
                c: this._colIndex
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
                    oncell.call(thisArg, cell, index, this._colIndex);
            } else {
                row.push(cell);
            }
        }

        if (hasRecordCb) onrecord.call(thisArg, row, index);
        return row;
    }

    /**
     * Read count rows starting at the specified index.
     * @param {number} startIndex The relative index (negative, zero, positive)
     * at which to start reading. A negative value instructs to read from the end.
     * @param {number} count The maximum number of rows to read.
     * @returns {any[]} An one-dimensional array of arrays representing the rows read.
     */
    readMany(startIndex, count) {
        // force cast these numbers to integers
        startIndex = startIndex | 0;
        count = count | 0;

        if (count < 0) {
            this._handleError(new Error("count cannot be negative."));
            return null;
        }

        const rows = [];
        let counter = -1;

        if (startIndex < 0) {
            // read backwards
            startIndex = this._endRow + startIndex;
            while (++counter < count && (startIndex - counter > -1)) {
                const row = this.read(this._rowIndex = startIndex - counter);
                rows.push(row);
            }
        } else {
            while (++counter < count && (startIndex + counter <= this._endRow)) {
                const row = this.read(this._rowIndex = startIndex + counter);
                rows.push(row);
            }
        }

        return rows;
    }

    /**
     * Read and return a cell at the specified row and column indices.
     * @param {number} colIndex The zero-based column index to read.
     * @param {number} rowIndex The zero-based row index to read.
     * @returns {{t: string, v: any, r: string, h: any, w: string}}
     * An object that represents the cell at the specified indices.
     */
    readCell(colIndex, rowIndex) {
        rowIndex || (rowIndex = this._rowIndex || this._startRow);
        colIndex || (colIndex = this._colIndex || this._startCol);

        const encodedCell = xlsx.utils.encode_cell({
            r: rowIndex,
            c: colIndex
        });

        return this._sheet[encodedCell];
    }

    /**
     * Register an event handler.
     * @param {"start"|"cell"|"record"|"end"|"error"} eventName The name of the event to register.
     * @param {Function} handler A callback function that handles the specified event.
     */
    on(eventName, handler) {
        if (SUPPORTED_EVENTS.indexOf(eventName) === -1)
            throw new Error("Unknown event: " + eventName);

        this["on" + eventName] = handler;
        return this;
    }

    /**
     * Destroy the internal WorkSheet.
     */
    destroy() {
        if (!this._destroyed) {
            delete this._sheet;
            this._sheet = null;
            this._destroyed = true;
        }
    }

    /**
     * Fire the 'error' event handler (if any), or throw error.
     * @param {Error|string} error The error to report.
     */
    _handleError(error) {
        if (typeof error === "string")
            error = new Error(error);

        if (!!this.onerror) {
            this.onerror.call(this._thisArg, error);
        } else {
            throw error;
        }
    }

    _isDestroyed() {
        if (this._destroyed || this._sheet === null) {
            this._handleError("The WorkSheet reader has been destroyed.");
            return true;
        }
        return false;
    }
}

const SUPPORTED_EVENTS = ["start", "cell", "beforerecord", "record", "end", "error"];

/**
 * Make sure that the specified value is a function.
 * @param {Function} value A callback function.
 */
const _ensureFunction = value => {
    if (value !== undefined && value !== null && typeof value !== "function") {
        throw new Error("value is not a function.");
    }
    return value;
}

module.exports = FastXlsxSheetReader;