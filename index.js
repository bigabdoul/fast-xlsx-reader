const FastXlsxReader = require("./lib/FastXlsxReader");

exports = module.exports = FastXlsxReader;

/**
 * Read sequentially the rows contained in an Excel sheet.
 * @param {{
      input: string,
      output?: string|WriteStream|null,
      format?: string,
      sheetname?: string|number,
      hasHeader?: boolean,
      headerPrefix?: string,
      lowerCaseHeaders?: boolean,
      schema?: object,
      onHeader?: (header: string[]) => void,
      onCell?: (cell, rowIndex: number, colIndex: number) => void,
      onRecord?: (record: any[], index: number) => void,
      onFinish?: (items: any[], rowsProcessed: number) => void,
      onError?: (err) => void,
      useMemoryForItems?: boolean,
      backwards?: boolean
    }} options An object containing processing instructions.
 */
exports.read = options => new FastXlsxReader(options).read();

/**
 * Create and return an instance of the FastXlsxSheetReader class.
 * @param {{
      input: string,
      output?: string|WriteStream|null,
      format?: string,
      sheetname?: string|number,
      hasHeader?: boolean,
      headerPrefix?: string,
      lowerCaseHeaders?: boolean,
      schema?: object,
      onHeader?: (header: string[]) => void,
      onCell?: (cell, rowIndex: number, colIndex: number) => void,
      onRecord?: (record: any[], index: number) => void,
      onFinish?: (items: any[], rowsProcessed: number) => void,
      onError?: (err) => void,
      useMemoryForItems?: boolean,
      backwards?: boolean
    }} options An object containing processing instructions.
 */
exports.createReader = options => new FastXlsxReader(options).createReader();