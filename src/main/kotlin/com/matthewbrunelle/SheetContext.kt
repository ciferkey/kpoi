package com.matthewbrunelle

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress

fun Sheet.row(rowNumber: Int? = null, block: Row.() -> Unit): Row {
    return createRow(rowNumber ?: physicalNumberOfRows).apply(block)
}

fun Sheet.merge(firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int) {
    addMergedRegion(CellRangeAddress(firstRow, lastRow, firstCol, lastCol))
}

fun Sheet.style(block: CellStyle.() -> Unit): CellStyle {
    // Note, unlock at the cell and row level this style will not be set on anything
    return workbook.createCellStyle().apply(block)
}