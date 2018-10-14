package com.matthewbrunelle

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle

/**
 * Returns given a receiver Cell creates a new style for the cell and applies the block to the style.
 */
fun Cell.style(block: CellStyle.() -> Unit): CellStyle {
    // TODO: there is an upper limit to the number of styles a workbook can contain, leverage CellUtil to avoid repeating styles
    val s = sheet.workbook.createCellStyle().apply(block)
    cellStyle = s
    return s
}