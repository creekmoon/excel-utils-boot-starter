package cn.creekmoon.excel.support.assertion;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** 基于 Apache POI 打开导出/回写文件并做结构断言（与 Hutool 写入解耦）。 */
public final class PoiWorkbookAssertions {

    private PoiWorkbookAssertions() {
    }

    public static Workbook open(File file) throws IOException {
        try (FileInputStream inputStream = new FileInputStream(file)) {
            return WorkbookFactory.create(inputStream);
        }
    }

    public static Workbook open(InputStream inputStream) throws IOException {
        return WorkbookFactory.create(inputStream);
    }

    public static void assertSheetNames(Workbook workbook, String... names) {
        assertEquals(names.length, workbook.getNumberOfSheets());
        for (int i = 0; i < names.length; i++) {
            assertEquals(names[i], workbook.getSheetName(i));
        }
    }

    public static void assertRowCount(Sheet sheet, int expected) {
        assertEquals(expected, countPhysicalRows(sheet));
    }

    public static int countPhysicalRows(Sheet sheet) {
        int maxRow = sheet.getLastRowNum();
        if (maxRow < 0) {
            return 0;
        }
        return maxRow + 1;
    }

    /** 按字符串比较单元格（数值格会先转成整数字符串）。 */
    public static void assertCellString(Sheet sheet, int rowIndex, int colIndex, String expected) {
        assertEquals(expected, readCellAsString(sheet, rowIndex, colIndex));
    }

    public static void assertCellValue(Sheet sheet, int rowIndex, int colIndex, String expected) {
        assertCellString(sheet, rowIndex, colIndex, expected);
    }

    public static void assertCellBlank(Sheet sheet, int rowIndex, int colIndex) {
        Cell cell = sheet.getRow(rowIndex) == null ? null : sheet.getRow(rowIndex).getCell(colIndex);
        if (cell == null) {
            return;
        }
        assertTrue(cell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK
                || cell.getStringCellValue().isBlank());
    }

    public static int findHeaderColumnIndex(Sheet sheet, int titleRowIndex, String header) {
        Row row = sheet.getRow(titleRowIndex);
        assertNotNull(row, "标题行不存在: rowIndex=" + titleRowIndex);
        short lastCellNum = row.getLastCellNum();
        for (int col = 0; col < lastCellNum; col++) {
            Cell cell = row.getCell(col);
            if (cell != null && header.equals(cell.getStringCellValue())) {
                return col;
            }
        }
        throw new AssertionError("未找到表头列: " + header);
    }

    public static void assertMergedRegionsCount(Sheet sheet, int expected) {
        assertEquals(expected, sheet.getNumMergedRegions());
    }

    public static void assertCellFillForeground(Sheet sheet, int rowIndex, int colIndex, IndexedColors expected) {
        Cell cell = requireCell(sheet, rowIndex, colIndex);
        short color = cell.getCellStyle().getFillForegroundColor();
        assertEquals(expected.getIndex(), color,
                "单元格(" + rowIndex + "," + colIndex + ")前景色不符合预期");
    }

    /** 读取「导入结果」列：行号 → 回写消息（跳过空白格）。 */
    public static Map<Integer, String> readResultColumn(Sheet sheet, int titleRowIndex, String resultTitle) {
        int resultCol = findHeaderColumnIndex(sheet, titleRowIndex, resultTitle);
        Map<Integer, String> rowIndex2Message = new HashMap<>();
        for (int row = titleRowIndex + 1; row <= sheet.getLastRowNum(); row++) {
            Row dataRow = sheet.getRow(row);
            if (dataRow == null) {
                continue;
            }
            Cell cell = dataRow.getCell(resultCol);
            if (cell == null || cell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK) {
                continue;
            }
            rowIndex2Message.put(row, cell.getStringCellValue());
        }
        return rowIndex2Message;
    }

    public static String readCellAsString(Sheet sheet, int rowIndex, int colIndex) {
        Cell cell = sheet.getRow(rowIndex) == null ? null : sheet.getRow(rowIndex).getCell(colIndex);
        if (cell == null) {
            return null;
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case BLANK -> null;
            default -> cell.toString();
        };
    }

    public static boolean containsMergedRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() == firstRow && region.getLastRow() == lastRow
                    && region.getFirstColumn() == firstCol && region.getLastColumn() == lastCol) {
                return true;
            }
        }
        return false;
    }

    private static Cell requireCell(Sheet sheet, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        assertNotNull(row, "行不存在: rowIndex=" + rowIndex);
        Cell cell = row.getCell(colIndex);
        assertNotNull(cell, "单元格不存在: rowIndex=" + rowIndex + ", colIndex=" + colIndex);
        return cell;
    }
}
