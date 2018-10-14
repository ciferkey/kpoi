package com.matthewbrunelle

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle

/**
 * Creates a new cellStyle for the given receiver Cell and applies the [block] to the style.
 */
fun Cell.style(block: CellStyle.() -> Unit): CellStyle {
    // TODO: there is an upper limit to the number of styles a workbook can contain, leverage CellUtil to avoid repeating styles
    return sheet.workbook.createCellStyle()
            .apply(block)
            .apply(this::setCellStyle)
}