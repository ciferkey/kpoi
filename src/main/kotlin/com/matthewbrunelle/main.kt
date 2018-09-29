package com.matthewbrunelle

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import java.io.FileInputStream
import java.io.FileOutputStream
import java.util.*


data class Sheet2(val sheet: Sheet, var rowCount: Int = 0) {
    fun addRow(): Row2 {
        val rowNum = rowCount
        rowCount += 1
        return Row2(sheet.createRow(rowNum)!!)
    }
}

data class Row2(val row: Row, var cellCount: Int = 0) {
    fun addCell(): Cell {
        val cellNum = cellCount
        cellCount += 1
        return row.createCell(cellNum)!!
    }
}

fun workbook(wb: Workbook = HSSFWorkbook(), block: Workbook.() -> Unit): Workbook {
    block(wb)
    return wb
}

fun Workbook.sheet(name: String? = null, block: Sheet2.() -> Unit): Sheet2 {
    // TODO: find idiomatic way
    val s = if (name != null) {
        createSheet(name)
    } else {
        createSheet()
    }
    return Sheet2(s).apply(block)
}

fun Workbook.write(fileName: String) {
    FileOutputStream(fileName).use { fileOut -> this.write(fileOut) }
}

fun WorkbookFactory.read(fileName: String) {
    FileInputStream(fileName).use { fileIn -> WorkbookFactory.create(fileIn) }
}

fun Sheet2.row(block: Row2.() -> Unit): Row2 {
    return addRow().apply(block)
}

fun Row2.cell(value: Any? = null, block: Cell.() -> Unit): Cell {
    // TODO: more idiomatic way?
    val cell = addCell()
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