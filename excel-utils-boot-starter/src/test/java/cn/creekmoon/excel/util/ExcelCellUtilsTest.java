package cn.creekmoon.excel.util;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * ExcelCellUtils 坐标换算契约。
 *
 * <p>注意：列/行解析函数返回 0-based 索引（"A"→0，"B6"→行5/列1），
 * 供 EasyExcelCellReader 等内部直接用作数组下标；excelIndexToCell 反向输出 1-based 单元格引用。</p>
 */
class ExcelCellUtilsTest {

    @Test
    void excelColumnToIndex_singleLetter() {
        assertEquals(0, ExcelCellUtils.excelColumnToIndex("A"));
        assertEquals(25, ExcelCellUtils.excelColumnToIndex("Z"));
    }

    @Test
    void excelColumnToIndex_multiLetter() {
        assertEquals(26, ExcelCellUtils.excelColumnToIndex("AA"));
        assertEquals(51, ExcelCellUtils.excelColumnToIndex("AZ"));
        assertEquals(52, ExcelCellUtils.excelColumnToIndex("BA"));
        assertEquals(702, ExcelCellUtils.excelColumnToIndex("AAA"));
    }

    @Test
    void excelCellToColumnIndex_extractsZeroBasedColumnPart() {
        assertEquals(1, ExcelCellUtils.excelCellToColumnIndex("B6"));
        assertEquals(26, ExcelCellUtils.excelCellToColumnIndex("AA23"));
        assertEquals(0, ExcelCellUtils.excelCellToColumnIndex("A1"));
    }

    @Test
    void excelCellToRowIndex_extractsZeroBasedRowPart() {
        assertEquals(5, ExcelCellUtils.excelCellToRowIndex("B6"));
        assertEquals(22, ExcelCellUtils.excelCellToRowIndex("AA23"));
        assertEquals(0, ExcelCellUtils.excelCellToRowIndex("A1"));
    }

    @Test
    void excelIndexToCell_zeroBasedIndicesToCellReference() {
        assertEquals("A1", ExcelCellUtils.excelIndexToCell(0, 0));
        assertEquals("A10", ExcelCellUtils.excelIndexToCell(9, 0));
        assertEquals("AA28", ExcelCellUtils.excelIndexToCell(27, 26));
    }

    @Test
    void coordinateRoundTrip_invertsExactly() {
        for (int row = 0; row < 200; row += 13) {
            for (int col = 0; col < 100; col += 7) {
                String cell = ExcelCellUtils.excelIndexToCell(row, col);
                assertEquals(row, ExcelCellUtils.excelCellToRowIndex(cell), "row of " + cell);
                assertEquals(col, ExcelCellUtils.excelCellToColumnIndex(cell), "col of " + cell);
            }
        }
    }
}
