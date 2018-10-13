package com.matthewbrunelle

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import java.io.FileInputStream
import java.io.FileOutputStream

fun workbook(wb: Workbook = HSSFWorkbook(), block: Workbook.() -> Unit): Workbook {
    block(wb)
    return wb
}

fun Workbook.sheet(name: String? = null, block: Sheet.() -> Unit): Sheet {
    return if (name != null) {
        createSheet(name)
    } else {
        createSheet()
    }.apply(block)
}

fun Workbook.style(block: CellStyle.() -> Unit): CellStyle {
    // Note, unlock at the cell and row level this style will not be set on anything
    return createCellStyle().apply(block)
}

fun Workbook.font(style: CellStyle, block: Font.() -> Unit): Font {
    val font = createFont().apply(block)
    style.setFont(font)
    return font
}

fun Workbook.write(fileName: String) {
    FileOutputStream(fileName).use { fileOut -> this.write(fileOut) }
}

fun WorkbookFactory.read(fileName: String) {
    FileInputStream(fileName).use { fileIn -> WorkbookFactory.create(fileIn) }
}