package cn.creekmoon.excel.support.fixture;

import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * 测试 fixture 元数据：启动时用 POI 预读行数、抽样行、单元格期望值，避免用例内硬编码。
 */
public final class FixtureCatalog {

    /** 标准表头导入（约 1000 行数据）。 */
    public static final String STANDARD_IMPORT = "import-demo-1000.xlsx";
    /** 单元格坐标读取专用（含合并区 sheet）。 */
    public static final String CELL_READ_IMPORT = "import-demo-cell-read.xlsx";

    private static int totalDataRows;
    private static SampleRow sampleRow6;
    private static SampleRow sampleRow80;
    private static SampleRow sampleRow195;
    private static SampleRow sampleRow631;
    private static long ageGreaterThan60Count;
    private static int sheet1DataRows;
    private static Map<String, String> cellSheet0Expectations;
    private static Map<String, String> cellSheet1Expectations;

    static {
        loadStandardImport();
        loadCellReadImport();
    }

    private FixtureCatalog() {
    }

    public static int totalDataRows() {
        return totalDataRows;
    }

    public static SampleRow sampleRow(int excelRowNumber) {
        return switch (excelRowNumber) {
            case 6 -> sampleRow6;
            case 80 -> sampleRow80;
            case 195 -> sampleRow195;
            case 631 -> sampleRow631;
            default -> throw new IllegalArgumentException("未配置样本行: " + excelRowNumber);
        };
    }

    public static long ageGreaterThan60Count() {
        return ageGreaterThan60Count;
    }

    public static int sheet1DataRows() {
        return sheet1DataRows;
    }

    public static Map<String, String> cellSheet0Expectations() {
        return cellSheet0Expectations;
    }

    public static Map<String, String> cellSheet1Expectations() {
        return cellSheet1Expectations;
    }

    private static void loadStandardImport() {
        try (InputStream inputStream = FixtureCatalog.class.getResourceAsStream("/fixtures/import/" + STANDARD_IMPORT);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<String, Integer> headerIndex = readHeaderIndex(sheet.getRow(0));
            int dataRows = 0;
            long ageCount = 0;
            SampleRow row6 = null;
            SampleRow row80 = null;
            SampleRow row195 = null;
            SampleRow row631 = null;
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }
                dataRows++;
                int excelRowNumber = rowIndex + 1;
                SampleRow sample = toSampleRow(row, headerIndex, excelRowNumber);
                if (excelRowNumber == 6) {
                    row6 = sample;
                } else if (excelRowNumber == 80) {
                    row80 = sample;
                } else if (excelRowNumber == 195) {
                    row195 = sample;
                } else if (excelRowNumber == 631) {
                    row631 = sample;
                }
                if (sample.age() != null && sample.age() > 60) {
                    ageCount++;
                }
            }
            totalDataRows = dataRows;
            sampleRow6 = row6;
            sampleRow80 = row80;
            sampleRow195 = row195;
            sampleRow631 = row631;
            ageGreaterThan60Count = ageCount;

            Sheet sheet1 = workbook.getSheetAt(1);
            sheet1DataRows = Math.max(0, sheet1.getLastRowNum());
        } catch (IOException e) {
            throw new IllegalStateException("加载标准导入 fixture 失败", e);
        }
    }

    private static void loadCellReadImport() {
        cellSheet0Expectations = Map.of(
                "B1", "李二狗",
                "D1", "李云龙",
                "age", "2000"
        );
        cellSheet1Expectations = Map.of(
                "A1", "合并单元格(A1)",
                "E1", "合并单元格(E1)",
                "A5", "合并单元格(A5)",
                "E3", "合并单元格(E3)"
        );
    }

    private static Map<String, Integer> readHeaderIndex(Row headerRow) {
        Map<String, Integer> headerIndex = new HashMap<>();
        for (int col = 0; col < headerRow.getLastCellNum(); col++) {
            if (headerRow.getCell(col) != null) {
                headerIndex.put(headerRow.getCell(col).getStringCellValue(), col);
            }
        }
        return headerIndex;
    }

    private static SampleRow toSampleRow(Row row, Map<String, Integer> headerIndex, int excelRowNumber) {
        return new SampleRow(
                excelRowNumber,
                readCell(row, headerIndex.get("用户名")),
                readCell(row, headerIndex.get("全名")),
                readCell(row, headerIndex.get("邮箱")),
                parseAge(readCell(row, headerIndex.get("年龄")))
        );
    }

    private static String readCell(Row row, Integer colIndex) {
        if (colIndex == null || row.getCell(colIndex) == null) {
            return null;
        }
        return row.getCell(colIndex).toString();
    }

    private static Integer parseAge(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        try {
            return IntegerConverter.parse(value);
        } catch (cn.creekmoon.excel.util.exception.CheckedExcelException e) {
            return null;
        }
    }

    public record SampleRow(int excelRowNumber, String userName, String fullName, String email, Integer age) {
    }
}
