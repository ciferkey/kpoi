package com.matthewbrunelle

import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet

fun Sheet.row(rowNumber: Int? = null, block: Row.() -> Unit): Row {
    return createRow(rowNumber ?: physicalNumberOfRows).apply(block)
}