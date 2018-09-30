package com.matthewbrunelle

import org.apache.poi.ss.usermodel.*

fun CellStyle.align(horizontalAlignment: HorizontalAlignment, verticalAlignment: VerticalAlignment) {
    alignment = horizontalAlignment
    this.verticalAlignment = verticalAlignment
}