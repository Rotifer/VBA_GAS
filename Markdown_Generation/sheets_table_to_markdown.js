/**
 * Helper function to create the table Markdown hyphen row that sits
 *   between the column names and the data rows of the table and determines the table cell alignment
 * 
 * @param {string} alignment - One of "left", "mid" or "right" to specify cell alignment in the Markdown table
 * @param {number} columnCount - The number of columns in the returned Markdown table.
 * @return Text containing the Markdown table hyphen line specifying the table cell alignment foor each table column.
 */
function getHyphenLine(alignment, columnCount) {
  if (alignment == "left") {
    return `|${Array(columnCount).fill(":---").join("|")}|`;
  } else if (alignment == "right") {
    return `|${Array(columnCount).fill("---:").join("|")}|`;
  } else {
    return `|${Array(columnCount).fill(":---:").join("|")}|`;
  }
}

/**
 * Create Markdown text from an input spreadsheet range.
 *
 * @param {string[]} rngValues - The range input as A1:B2 syntax.
 * @param {string} alignment - The default is "mid" for the alignment.
 * @return The input range values as table Markdown text.
 * @customfunction
 */
function TABLE_TO_MARKDOWN(rngValues, alignment = "mid") {
  const columnNames = rngValues[0];
  const tableRows = rngValues.slice(1);
  const hyphenLine = getHyphenLine(alignment, columnNames.length);
  const headerHyphenLines = `|${columnNames.join("|")}|\n${hyphenLine}`;
  const tableRowLines = tableRows.map((tableRow) => {
    return `|${tableRow.join("|")}|`;
  }).join("\n");
  return `${headerHyphenLines}\n${tableRowLines}`;
}
