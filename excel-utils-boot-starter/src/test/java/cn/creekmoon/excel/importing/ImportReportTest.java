package cn.creekmoon.excel.importing;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.report.ImportGlobalErrorEnum;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.assertion.PoiWorkbookAssertions;
import cn.creekmoon.excel.support.fixture.FixtureCatalog;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.XlsxTestBuilder;
import cn.creekmoon.excel.support.spring.TestBizException;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.ExcelFileUtils;
import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.creekmoon.excel.util.exception.CustomExceptionHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** ImportReport：统计、全局错误、行级错误明细、业务异常策略。 */
@Slf4j
class ImportReportTest {

    /** 注册测试用业务异常处理器，供自定义异常相关用例使用。 */
    @BeforeAll
    static void registerTestBizException() {
        ExcelUtilsConfig.addExcelUtilsExceptionHandler(new CustomExceptionHandler(TestBizException.class));
        log.info("[SETUP] registered TestBizException handler");
    }

    @Nested
    class Smoke {

        /** 无配置读取：统计完整、无全局错误、报告不携带业务对象。 */
        @Test
        void importReport_defaultBehavior_returnsStructuredStatistics() throws Exception {
            log.info("[STEP] readWithReport without any options");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);

            ImportReport report = ImportTestHelper.bindStandardColumns(
                            excelImport.switchSheet(0, Student::new))
                    .readWithReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(FixtureCatalog.totalDataRows(), report.getSuccessRows());
            assertEquals(FixtureCatalog.totalDataRows(), report.getTotalRows());
            assertEquals(0, report.getErrorRows());
            assertTrue(report.getGlobalErrors().isEmpty());
            ImportReportAssertions.assertBalancedStatistics(report);
            log.info("[DONE] default report passed");
        }
    }

    /** 分支：模板、全表失败、异常包装等。 */
    @Nested
    class Branch {

        /** 模板列缺失：立即停止，不扫描任何数据行。 */
        @Test
        void importReport_templateMismatch_stopsImmediately() throws Exception {
            log.info("[STEP] configure missing title column");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);

            ImportReport report = excelImport.switchSheet(0, Student::new)
                    .addConvert("missing-title-column", Student::setUserName)
                    .readWithReport();
            ImportReportAssertions.logReportSummary(report);

            assertTrue(report.getGlobalErrors().contains(ImportGlobalErrorEnum.TEMPLATE_MISMATCH));
            assertEquals(0, report.getTotalRows());
            assertEquals(0, report.getSuccessRows());
            assertEquals(0, report.getErrorRows());
            assertTrue(report.getRowErrors().isEmpty(), "模板不匹配属于文件级失败，不产生行级明细");
            log.info("[DONE] template mismatch fail-fast passed");
        }

        /** 每行后置校验都失败：读满全表，好行不因错误数量被截断。 */
        @Test
        void importReport_everyRowFails_stillScansWholeSheet() throws Exception {
            log.info("[STEP] inject CheckedExcelException on every row");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);

            ImportReport report = ImportTestHelper.bindStandardColumns(
                            excelImport.switchSheet(0, Student::new))
                    .addConvertPostProcessor(student -> {
                        throw new CheckedExcelException("forced row validation error");
                    })
                    .readWithReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(FixtureCatalog.totalDataRows(), report.getTotalRows());
            assertEquals(FixtureCatalog.totalDataRows(), report.getErrorRows());
            assertEquals(FixtureCatalog.totalDataRows(), report.getRowErrors().size());
            assertTrue(report.getRowErrors().get(0).getMessage().contains("[失败!]"));
            assertTrue(report.getGlobalErrors().isEmpty(), "行级错误不产生全局错误");
            ImportReportAssertions.assertBalancedStatistics(report);
            log.info("[DONE] full scan under all-failing rows passed");
        }

        /** 已注册的 TestBizException：记错行但不中断，后续行继续读。 */
        @Test
        void importReport_customBizException_doesNotInterruptRemainingRows() throws Exception {
            AtomicInteger visited = new AtomicInteger();
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportReport report = ImportTestHelper.bindStandardColumns(
                            excelImport.switchSheet(0, Student::new))
                    .read((rowIndex, student) -> {
                        visited.incrementAndGet();
                        if (rowIndex == 10) {
                            throw new TestBizException("business validation failed at row 11");
                        }
                    })
                    .getReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(FixtureCatalog.totalDataRows(), report.getTotalRows());
            assertEquals(999, report.getSuccessRows());
            assertEquals(1, report.getErrorRows());
            assertEquals(FixtureCatalog.totalDataRows(), visited.get());
            assertFalse(report.getGlobalErrors().contains(ImportGlobalErrorEnum.IMPORT_INTERRUPTED));
            assertTrue(report.getRowErrors().get(0).getMessage().contains("business validation failed"));
            ImportReportAssertions.assertBalancedStatistics(report);
            log.info("[DONE] custom biz exception non-interrupt passed");
        }

        /** 业务异常包在 RuntimeException 内：视为非自定义异常，导入中断。 */
        @Test
        void importReport_wrappedBizException_marksImportInterrupted() throws Exception {
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportReport report = ImportTestHelper.bindStandardColumns(
                            excelImport.switchSheet(0, Student::new))
                    .read((rowIndex, student) -> {
                        if (rowIndex == 10) {
                            throw new RuntimeException(new TestBizException("wrapped biz exception at row 11"));
                        }
                    })
                    .getReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(10, report.getTotalRows());
            assertEquals(9, report.getSuccessRows());
            assertEquals(1, report.getErrorRows());
            assertTrue(report.getGlobalErrors().contains(ImportGlobalErrorEnum.IMPORT_INTERRUPTED));
            ImportReportAssertions.assertBalancedStatistics(report);
            log.info("[DONE] wrapped biz exception interrupt passed");
        }

        /** 年龄>90 抛 CheckedExcelException：total = success + error，单条业务错误对应一条明细。 */
        @Test
        void importReport_statistics_totalEqualsSuccessPlusError() throws Exception {
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportReport report = ImportTestHelper.bindStandardColumns(
                            excelImport.switchSheet(0, Student::new))
                    .read(student -> {
                        if (student.getAge() != null && student.getAge() > 90) {
                            throw new CheckedExcelException("age > 90 rejected");
                        }
                    })
                    .getReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(FixtureCatalog.totalDataRows(), report.getTotalRows());
            ImportReportAssertions.assertBalancedStatistics(report);
            assertTrue(report.getErrorRows() > 0);
            assertEquals(report.getErrorRows(), report.getRowErrors().size());
            log.info("[DONE] statistics consistency passed, errorRows={}", report.getErrorRows());
        }
    }

    /** 单行多字段失败：一次报全，供填表人一轮改完。 */
    @Nested
    class MultiFieldError {

        private static final Object[][] TWO_BAD_COLUMNS = {
                {"用户名", "年龄", "生日"},
                {"user1", "abc", "not-a-date"},   // 年龄、生日两列同时失败
                {"user2", 30, "2024-01-01"},
        };

        /** 同一行两列转换失败：错误行数为 1，行级明细两条且各带字段名。 */
        @Test
        void multipleFieldFailures_inOneRow_reportedTogether() throws Exception {
            ExcelImport excelImport = ExcelImport.create(
                    XlsxTestBuilder.multipartFile("multi-error.xlsx", rows(TWO_BAD_COLUMNS)));
            ImportReport report = excelImport.switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("生日", DateConverter::parse, Student::setBirthday)
                    .readWithReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(1, report.getErrorRows(), "错误行数按行计");
            assertEquals(1, report.getSuccessRows());
            assertEquals(2, report.getRowErrors().size(), "两个失败字段各一条明细");
            assertTrue(report.getRowErrors().stream().allMatch(x -> x.getRowIndex() == 1));
            assertTrue(report.getRowErrors().stream().anyMatch(x -> "年龄".equals(x.getField())));
            assertTrue(report.getRowErrors().stream().anyMatch(x -> "生日".equals(x.getField())));
            assertTrue(report.getRowErrors().stream().allMatch(x -> x.getMessage().contains("[失败!]")));
            ImportReportAssertions.assertBalancedStatistics(report);
        }

        /** 必填缺失与转换失败并存：同一行两条明细，消息分别命中两类原因。 */
        @Test
        void mustExistAndConvertFailure_inOneRow_reportedTogether() throws Exception {
            Object[][] table = {
                    {"用户名", "年龄"},
                    {"", "abc"},                  // 用户名必填缺失 + 年龄转换失败
                    {"user2", 30},
            };
            ExcelImport excelImport = ExcelImport.create(
                    XlsxTestBuilder.multipartFile("mixed-error.xlsx", rows(table)));
            ImportReport report = excelImport.switchSheet(0, Student::new)
                    .addConvertAndMustExist("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .readWithReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(1, report.getErrorRows());
            assertEquals(2, report.getRowErrors().size());
            assertTrue(report.getRowErrors().stream()
                    .anyMatch(x -> x.getMessage().contains("字段[用户名]为必填项")));
            assertTrue(report.getRowErrors().stream()
                    .anyMatch(x -> x.getMessage().contains("字段[年龄]解析失败")));
        }

        /** 多条明细回写到结果列时合并为一格，以统一分隔符连接。 */
        @Test
        void multipleFieldFailures_mergedIntoSingleResultCell() throws Exception {
            ExcelImport excelImport = ExcelImport.create(
                    XlsxTestBuilder.multipartFile("multi-error-writeback.xlsx", rows(TWO_BAD_COLUMNS)));
            excelImport.switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("生日", DateConverter::parse, Student::setBirthday)
                    .readWithReport();

            File file = excelImport.generateResultFile(false);
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                Map<Integer, String> resultMessages = PoiWorkbookAssertions.readResultColumn(
                        sheet, 0, ExcelConstants.RESULT_TITLE);
                log.info("[RESULT] merged result column: {}", resultMessages);

                assertEquals(1, resultMessages.size(), "两条明细写在同一行的同一格");
                String message = resultMessages.get(1);
                assertNotNull(message);
                assertTrue(message.contains("字段[年龄]解析失败"));
                assertTrue(message.contains("字段[生日]解析失败"));
                assertTrue(message.contains(ExcelConstants.ROW_ERROR_MSG_SEPARATOR));
            }
            ExcelFileUtils.cleanTempFileByPathNow(file.getPath());
        }

        /** 后置处理器失败仍只记一条：跨字段校验存在依赖，不做多错收集。 */
        @Test
        void postProcessorFailure_staysSingleRowError() throws Exception {
            Object[][] table = {
                    {"用户名", "年龄"},
                    {"user1", 20},
                    {"user2", 30},
            };
            ExcelImport excelImport = ExcelImport.create(
                    XlsxTestBuilder.multipartFile("pp-single.xlsx", rows(table)));
            ImportReport report = excelImport.switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvertPostProcessor(student -> {
                        throw new CheckedExcelException("post processor rejected");
                    })
                    .readWithReport();

            assertEquals(2, report.getErrorRows());
            assertEquals(2, report.getRowErrors().size());
            assertTrue(report.getRowErrors().stream().allMatch(x -> x.getField() == null));
        }
    }

    /** 表头行 + 数据行构建器（Object[][] 转 List<Object[]>，不能直接用 List.of）。 */
    private static List<Object[]> rows(Object[][] table) {
        return Arrays.asList(table);
    }
}
