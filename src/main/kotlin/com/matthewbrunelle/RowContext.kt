package com.matthewbrunelle

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Row
import java.util.*

/**
 * Create a new Cell for the given receiver Row and applies the [block] to it.
 * Optionally sets the value for the cell if it is provided.
 * Optionally sets the column of the cell if it is provided.
 */
fun Row.cell(value: Any? = null, column: Int? = null, block: Cell.() -> Unit = {}): Cell {
    val cell = createCell(column ?: physicalNumberOfCells)
    when (value) {
        is Calendar -> cell.setCellValue(value)
        is Date -> cell.setCellValue(value)
        is Boolean -> cell.setCellValue(value)
        is Double -> cell.setCellValue(value)
        is String -> cell.setCellValue(value)
        is RichTextString -> cell.setCellValue(value)
    }
    return cell.apply(block)
}

/**
 * Creates a new CellStyle for the given receiver Row and applies the [block] to the style.
 */
fun Row.style(block: CellStyle.() -> Unit): CellStyle {
    return sheet.workbook.createCellStyle()
            .apply(block)
            .apply(this::setRowStyle)
}