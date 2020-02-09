/**
 * Convert an Excel serial number to an instance of Date.
 * @param {number} serialDate The Excel serial number to convert.
 * @param {boolean} epoch1904 true to add 1462 extra days; otherwise, false.
 */
const parseDate = (serialDate, epoch1904) => {
    if (epoch1904) serialDate += 1462;
    const daysBeforeUnixEpoch = 70 * 365 + 19;
    const hour = 60 * 60 * 1000;
    return new Date(Math.round((serialDate - daysBeforeUnixEpoch) * 24 * hour) + 12 * hour);
};

/**
 * Attempt to convert the specified value to an instance of Date.
 * If the attempt fails, the value is returned as is.
 * @param {number|string} value The value to convert.
 * @param {boolean} epoch1904 true to add 1462 extra days; otherwise, false.
 */
const tryConvertDate = (value, epoch1904) => {
    try {
        if (typeof value === "number") {
            return parseDate(value, epoch1904);
        } else if (typeof value === "string") {
            let converted = parseInt(value);
            if (isNaN(converted)) {
                return new Date(Date.parse(value));
            } else
                return parseDate(converted, epoch1904);
        } else {
            return parseDate(parseInt(value.toString()), epoch1904);
        }
    } catch {
        return value;
    }
};

module.exports.parseDate = parseDate;
module.exports.tryConvertDate = tryConvertDate;