package com.matthewbrunelle

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle

fun Cell.style(block: CellStyle.() -> Unit): CellStyle {
    val s = sheet.workbook.createCellStyle().apply(block)
    cellStyle = s
    return s
}