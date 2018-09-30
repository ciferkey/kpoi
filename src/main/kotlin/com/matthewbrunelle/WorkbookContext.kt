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
    // TODO: find idiomatic way
    return if (name != null) {
        createSheet(name)
    } else {
        createSheet()
    }.apply(block)
}

fun Workbook.font(block: Font.() -> Unit): Font {
    return createFont().apply(block)
}

fun Workbook.write(fileName: String) {
    FileOutputStream(fileName).use { fileOut -> this.write(fileOut) }
}

fun WorkbookFactory.read(fileName: String) {
    FileInputStream(fileName).use { fileIn -> WorkbookFactory.create(fileIn) }
}