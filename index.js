const FastXlsxReader = require("./lib/FastXlsxReader");

exports = module.exports = FastXlsxReader;
exports.read = options => new FastXlsxReader(options).read();