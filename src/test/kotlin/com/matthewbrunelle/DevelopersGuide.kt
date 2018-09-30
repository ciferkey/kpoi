package com.matthewbrunelle

import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.BorderStyle.*
import org.apache.poi.ss.usermodel.CellType.ERROR
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.HorizontalAlignment.*
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.IndexedColors.*
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Test
import org.unitils.reflectionassert.ReflectionAssert.assertReflectionEquals
import org.unitils.reflectionassert.ReflectionComparatorMode.LENIENT_DATES
import java.util.*


/**
 * Translation of the "Busy Developers' Guide to HSSF and XSSF Features" )https://poi.apache.org/components/spreadsheet/quick-guide.html#NewWorkbook) to use this library. Also set up as runnable tests that compare their output to the original poi example's output.
 */


class DevelopersGuide {

    @Test
    fun newWorkBook_1_1() {
        val expectedWb = GenerateTestInputs.newWorkBook_1_1()

        val wb = workbook { }

        assertReflectionEquals(expectedWb, wb)
    }

    @Test
    fun newWorkBook_1_2() {
        val expectedWb = GenerateTestInputs.newWorkBook_1_2()

        val wb = workbook(XSSFWorkbook()) { }

        assertReflectionEquals(expectedWb, wb, LENIENT_DATES)
    }

    @Test
    fun newSheet() {
        val expectedWb = GenerateTestInputs.newSheet()

        val wb = workbook {
            sheet("new sheet") {

            }
            sheet("second sheet") {

            }
            // TODO: make all name safe?
            val safeName = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]")
            sheet(safeName) {

            }
        }

        assertReflectionEquals(expectedWb, wb)
    }

    @Test
    fun creatingCells() {
        val expectedWb = GenerateTestInputs.creatingCells()

        // TODO: extensions for creation helper?

        val wb = workbook {
            sheet("new sheet") {
                row {
                    // TODO: make lambda for cell optional?
                    cell(1.0) {}
                    cell(1.2) {}
                    cell(creationHelper.createRichTextString("This is a string")) { }
                    cell(true) {}
                }
            }
        }

        assertReflectionEquals(expectedWb, wb)
    }

    @Test
    fun creatingDateCells() {
//        AssertionsForInterfaceTypes.setLenientDateParsing(true)

        val expectedWb = GenerateTestInputs.creatingDateCells()

        val wb = workbook {
            // TODO: should you be able to do this in a inner scope?
            // TODO: mechanisms for defaulting cell styles?
            val dateCellStyle = createCellStyle()
            sheet("new sheet") {
                row {
                    // TODO: make lambda for cell optional?
                    cell(Date()) {}
                    cell(Date()) {
                        cellStyle = dateCellStyle
                    }
                    cell(Calendar.getInstance()) {
                        cellStyle = dateCellStyle
                    }
                }
            }
        }

        assertReflectionEquals(expectedWb, wb, LENIENT_DATES)
//        assertThat(expectedWb as Any)
//                .usingComparatorForType(RecursiveFieldByFieldComparator(emptyMap(), TypeComparators()), Workbook::class.java)
//                .usingComparatorForType(RecursiveFieldByFieldComparator(emptyMap(), TypeComparators()), Sheet::class.java)
//                .usingComparatorForType(RecursiveFieldByFieldComparator(emptyMap(), TypeComparators()), Row::class.java)
//                .usingComparatorForType(RecursiveFieldByFieldComparator(emptyMap(), TypeComparators()), Cell::class.java)
//                .isEqualToComparingFieldByFieldRecursively(wb)
    }

    @Test
    fun differentKindsOfCells() {
        val expectedWb = GenerateTestInputs.differentKindsOfCells()

        val wb = workbook {
            // TODO: should you be able to do this in a inner scope?
            // TODO: mechanisms for defaulting cell styles?
            sheet("new sheet") {
                row {
                    // TODO: make lambda for cell optional?
                    cell(1.1) {}
                    cell(Date()) {}
                    cell(Calendar.getInstance()) {}
                    cell("a string") {}
                    cell(true) {}
                    cell(ERROR) {}
                }
            }
        }

        assertReflectionEquals(expectedWb, wb, LENIENT_DATES)
    }

    @Test
    fun alignmentOptions() {
        val expectedWb = GenerateTestInputs.alignmentOptions()

        val wb = workbook {
            sheet {
                row(2) {
                    heightInPoints = 30F
                    cell("Align It") {
                        style {
                            align(HorizontalAlignment.CENTER, BOTTOM)
                        }
                    }
                    cell("Align It") {
                        style { align(CENTER_SELECTION, BOTTOM) }
                    }
                    cell("Align It") {
                        style { align(FILL, VerticalAlignment.CENTER) }
                    }
                    cell("Align It") {
                        style { align(GENERAL, VerticalAlignment.CENTER) }
                    }
                    cell("Align It") {
                        style { align(HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY) }
                    }
                    cell("Align It") {
                        style { align(LEFT, TOP) }
                    }
                    cell("Align It") {
                        style { align(RIGHT, TOP) }
                    }
                }
            }
        }

        assertReflectionEquals(expectedWb, wb, LENIENT_DATES)
    }

    @Test
    fun workingWithBorders() {
        val expectedWb = GenerateTestInputs.workingWithBorders()

        val wb = workbook {
            sheet("new sheet") {
                row(1) {
                    cell(4.0, 1) {
                        style {
                            borderBottom = THIN
                            bottomBorderColor = BLACK.index
                            borderBottom = THIN
                            bottomBorderColor = GREEN.index
                            borderBottom = THIN
                            bottomBorderColor = BLUE.index
                            borderBottom = MEDIUM_DASHED
                            bottomBorderColor = BLACK.index
                        }
                    }
                }
            }
        }

        assertReflectionEquals(expectedWb, wb, LENIENT_DATES)
    }

    @Test
    fun fillsAndColors() {
        val expectedWb = GenerateTestInputs.fillsAndColors()

        val wb = workbook {
            sheet("new sheet") {
                row(1) {
                    cell("X", 1) {
                        style {
                            fillBackgroundColor = IndexedColors.AQUA.getIndex()
                            fillPattern = FillPatternType.BIG_SPOTS
                        }
                    }
                    cell("X", 2) {
                        style {
                            fillForegroundColor = IndexedColors.ORANGE.getIndex()
                            fillPattern = FillPatternType.SOLID_FOREGROUND
                        }
                    }
                }
            }
        }

        assertReflectionEquals(expectedWb, wb, LENIENT_DATES)
    }

    @Test
    fun mergingCells() {
        val expectedWb = GenerateTestInputs.mergingCells()

        val wb = workbook {
            sheet("new sheet") {
                row(1) {
                    cell("This is a test of merging", 1) {}
                    merge(
                            1, //first row (0-based)
                            1, //last row  (0-based)
                            1, //first column (0-based)
                            2  //last column  (0-based)
                    )
                }
            }
        }

        assertReflectionEquals(expectedWb, wb, LENIENT_DATES)
    }

    @Test
    fun workingWithFonts() {
        val expectedWb = GenerateTestInputs.workingWithFonts()

        val wb = workbook {
            sheet("new sheet") {
                row(1) {
                    cell("This is a test of fonts", 1) {
                        style {
                            // TODO: why can't property notation be used here?
                            setFont(font {
                                fontHeightInPoints = 24.toShort()
                                fontName = "Courier New"
                                italic = true
                                strikeout = true
                            })
                        }
                    }
                }
            }
        }

        assertReflectionEquals(expectedWb, wb, LENIENT_DATES)
    }

    @Test
    fun customColors() {
        val expectedWb = GenerateTestInputs.customColors()

        val wb = workbook {
            sheet {
                row {
                    cell("Default Palette") {
                        style {
                            fillForegroundColor = HSSFColor.HSSFColorPredefined.LIME.index
                            fillPattern = FillPatternType.SOLID_FOREGROUND
                            setFont(font {
                                color = HSSFColor.HSSFColorPredefined.RED.index
                            })
                        }
                    }
                }
            }
        }

        assertReflectionEquals(expectedWb, wb, LENIENT_DATES)
    }
}
