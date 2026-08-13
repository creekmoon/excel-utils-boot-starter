package cn.creekmoon.excel.integration.example;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

/**
 * 复刻 excel-utils-example 的 20 万行大文件导入（fixture=导入测试20万行.xlsx）。
 *
 * <p>20 万行全量流式读取 + 抽样断言：只收集指定行号的样本（内存友好），
 * 样本值与源文件 POI 预读的对应行一致。</p>
 *
 * <p>文件布局：Sheet0 多级表头 2 行（row0 分组行、row1 子表头），数据 2..200001 共 200000 行；
 * Sheet1 单级表头 1 行，数据 1..10000 共 10000 行。</p>
 */
class ExampleLargeImportTest {

    private static final String FIXTURE = "导入测试20万行.xlsx";
    private static final int SHEET0_ROWS = 200_000;
    private static final int SHEET1_ROWS = 10_000;
    /** 抽样行下标（0-based，与 EasyExcel rowIndex / POI row 索引一致；2 为 Sheet0 首条数据行）。 */
    private static final int[] SAMPLE_ROW_INDEXES = {2, 50, 1000, 50000, 100000, 200001};

    @Test
    void largeImport_sheet0_twoHundredThousandRows_sampledAssertions() throws Exception {
        ExcelImport excelImport = ImportTestHelper.createImport(FIXTURE);

        /* 抽样收集器：rowIndex → 导入对象 */
        Map<Integer, Student> samples = new HashMap<>();

        TitleReader<Student> reader = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .range(1, SHEET0_ROWS + 1);
        reader.read((rowIndex, s) -> {
            for (int sampleRow : SAMPLE_ROW_INDEXES) {
                if (rowIndex == sampleRow) {
                    samples.put(rowIndex, s);
                }
            }
        });

        ImportReport report = reader.getReport();
        ImportReportAssertions.logReportSummary(report);
        assertEquals(SHEET0_ROWS, report.getTotalRows(), "Sheet0 应读到 20 万行");
        assertEquals(SHEET0_ROWS, report.getSuccessRows());
        ImportReportAssertions.assertNoErrors(report);

        /* 抽样断言：导入样本与源文件 POI 预读一致 */
        try (InputStream in = ImportTestHelper.class.getResourceAsStream("/fixtures/import/" + FIXTURE);
             Workbook workbook = WorkbookFactory.create(in)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int sampleRow : SAMPLE_ROW_INDEXES) {
                Row src = sheet.getRow(sampleRow);
                Student actual = samples.get(sampleRow);
                assertNotNull(actual, "抽样行 " + sampleRow + " 应被收集");
                assertEquals(src.getCell(0).getStringCellValue(), actual.getUserName(),
                        "抽样行 " + sampleRow + " 用户名");
                assertEquals(Double.valueOf(src.getCell(2).getNumericCellValue()).intValue(), actual.getAge(),
                        "抽样行 " + sampleRow + " 年龄");
                assertEquals(src.getCell(3).getStringCellValue(), actual.getEmail(),
                        "抽样行 " + sampleRow + " 邮箱");
                assertNotNull(actual.getBirthday(), "抽样行 " + sampleRow + " 生日应解析");
                assertNotNull(actual.getExpTime(), "抽样行 " + sampleRow + " 过期时间应解析");
            }
        }
    }

    @Test
    void largeImport_sheet1_tenThousandRows_sampledAssertions() throws Exception {
        ExcelImport excelImport = ImportTestHelper.createImport(FIXTURE);

        Map<Integer, Student> samples = new HashMap<>();
        int[] sampleRows = {1, 500, 5000, 10000};

        TitleReader<Student> reader = excelImport.switchSheet(1, Student::new)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .range(0, SHEET1_ROWS);
        reader.read((rowIndex, s) -> {
            for (int sampleRow : sampleRows) {
                if (rowIndex == sampleRow) {
                    samples.put(rowIndex, s);
                }
            }
        });

        ImportReport report = reader.getReport();
        ImportReportAssertions.logReportSummary(report);
        assertEquals(SHEET1_ROWS, report.getTotalRows(), "Sheet1 应读到 1 万行");
        assertEquals(SHEET1_ROWS, report.getSuccessRows());
        ImportReportAssertions.assertNoErrors(report);

        try (InputStream in = ImportTestHelper.class.getResourceAsStream("/fixtures/import/" + FIXTURE);
             Workbook workbook = WorkbookFactory.create(in)) {
            Sheet sheet = workbook.getSheetAt(1);
            for (int sampleRow : sampleRows) {
                Row src = sheet.getRow(sampleRow);
                Student actual = samples.get(sampleRow);
                assertNotNull(actual, "抽样行 " + sampleRow + " 应被收集");
                assertEquals(Double.valueOf(src.getCell(0).getNumericCellValue()).intValue(), actual.getAge(),
                        "抽样行 " + sampleRow + " 年龄");
                assertEquals(src.getCell(1).getStringCellValue(), actual.getEmail(),
                        "抽样行 " + sampleRow + " 邮箱");
            }
        }
    }
}
