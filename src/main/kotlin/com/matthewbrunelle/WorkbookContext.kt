package com.matthewbrunelle

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.WorkbookUtil
import java.io.FileInputStream
import java.io.FileOutputStream

/**
 * Creates a workbook if none is supplied and applies the [block] to it.
 * If a [workbook] is supplied it will be used instead.
 */
fun workbook(workbook: Workbook = HSSFWorkbook(), block: Workbook.() -> Unit = {}): Workbook {
    return workbook.apply(block)
}

/**
 * Creates a new Sheet in the receiver Workbook and applies the [block] to it.
 * If a [name] is provided it will be used for the sheet's name.
 */
fun Workbook.sheet(name: String? = null, block: Sheet.() -> Unit = {}): Sheet {
    return if (name != null) {
        val safeName = WorkbookUtil.createSafeSheetName(name)
        createSheet(safeName)
    } else {
        createSheet()
    }.apply(block)
}

/**
 * Creates a new CellStyle for the given receiver Workbook and applies the [block] to the style.
 * Note, unlike at the cell and row level this style will not be set on anything.
 */
fun Workbook.style(block: CellStyle.() -> Unit): CellStyle {
    return createCellStyle().apply(block)
}

/**
 * Creates a new font, applies the [block] to it and sets it on the given style.
 */
fun Workbook.font(style: CellStyle, block: Font.() -> Unit): Font {
    return createFont()
            .apply(block)
            .apply {
                style.setFont(this)
            }
}

/**
 * Convenience method for creating a RichTextString for a given piece of [text].
 */
fun Workbook.richText(text: String): RichTextString {
    return creationHelper.createRichTextString(text)
}

/**
 * Convenience method for creating a DateFormat on a Workbook for a given [dateFormat].
 */
fun Workbook.dateFormat(dateFormat: String): Short {
    return creationHelper.createDataFormat().getFormat(dateFormat)
}

/**
 * Convenience method for writing a Workbook to file with the given [fileName]
 */
fun Workbook.write(fileName: String) {
    FileOutputStream(fileName).use { fileOut -> this.write(fileOut) }
}

/**
 * Convenience method for reading a Workbook from a file with the given [fileName]
 */
fun WorkbookFactory.read(fileName: String) {
    FileInputStream(fileName).use { fileIn -> WorkbookFactory.create(fileIn) }
}