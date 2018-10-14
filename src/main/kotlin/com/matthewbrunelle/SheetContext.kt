package com.matthewbrunelle

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress

/**
 * Creates a new row for the given receiver Sheet and applies the [block] to the style.
 * Optionally sets the row number for the Row if its provided.
 */
fun Sheet.row(rowNumber: Int? = null, block: Row.() -> Unit): Row {
    return createRow(rowNumber ?: physicalNumberOfRows).apply(block)
}

/**
 * Convenience method for merging a region of cells in a Sheet.
 */
fun Sheet.merge(firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int) {
    addMergedRegion(CellRangeAddress(firstRow, lastRow, firstCol, lastCol))
}

/**
 * Creates a new CellStyle for the given receiver Row and applies the [block] to the style.
 * Note, unlike at the cell and row level this style will not be set on anything.
 */
fun Sheet.style(block: CellStyle.() -> Unit): CellStyle {
    return workbook.createCellStyle().apply(block)
}