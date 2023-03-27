function regexForStringSearch(query) {
    query = query
    .replace(/a/gi, '[AÁÀÂÃ]')
    .replace(/e/gi, '[EÉÈÊ]')
    .replace(/i/gi, '[IÍÌÎ]')
    .replace(/o/gi, '[OÓÒÔÕ]')
    .replace(/u/gi, '[UÚÙÛ]')
    .replace(/c/gi, '[CÇ]')
    return new RegExp(query, 'gmi')
}
module.exports = regexForStringSearch;