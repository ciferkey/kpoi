package com.matthewbrunelle

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Row
import java.util.*

fun Row.cell(value: Any? = null, column: Int? = null, block: Cell.() -> Unit = {}): Cell {
    // TODO: more idiomatic way?
    val cell = createCell(column ?: physicalNumberOfCells)
    value.let {
        when(value) {
            is Calendar -> cell.setCellValue(value)
            is Date -> cell.setCellValue(value)
            is Boolean -> cell.setCellValue(value)
            is Double -> cell.setCellValue(value)
            is String -> cell.setCellValue(value)
            is RichTextString -> cell.setCellValue(value)
        }
    }
    return cell.apply(block)
}

fun Row.style(block: CellStyle.() -> Unit): CellStyle {
    val s = sheet.workbook.createCellStyle().apply(block)
    rowStyle = s
    return s
}