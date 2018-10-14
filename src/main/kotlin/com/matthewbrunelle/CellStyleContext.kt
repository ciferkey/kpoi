package com.matthewbrunelle

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment

/**
 * applies a given [horizontalAlignment] and [verticalAlignment] to a CellStyle.
 */
fun CellStyle.align(horizontalAlignment: HorizontalAlignment, verticalAlignment: VerticalAlignment) {
    this.alignment = horizontalAlignment
    this.verticalAlignment = verticalAlignment
}