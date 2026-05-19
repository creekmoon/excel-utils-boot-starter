package cn.creekmoon.excel.importing;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.report.ImportGlobalErrorEnum;
import cn.creekmoon.excel.core.R.report.ImportOptions;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.fixture.FixtureCatalog;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.spring.TestBizException;
import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.creekmoon.excel.util.exception.CustomExceptionHandler;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** ImportReport / ImportOptions：统计、全局错误、错误阈值、业务异常策略。 */
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

        /** forApi 预设：fail-fast 模板、不收集成功数据、统计完整。 */
        @Test
        void importReport_apiPreset_returnsStructuredResultWithoutSuccessData() throws Exception {
            log.info("[STEP] readWithReport using ImportOptions.forApi()");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportOptions options = ImportOptions.forApi();
            options.setCollectRowErrors(true);

            ImportReport<Student> report = ImportTestHelper.bindStandardColumns(
                            excelImport.switchSheet(0, Student::new))
                    .readWithReport(options);
            ImportReportAssertions.logReportSummary(report);

            assertEquals(FixtureCatalog.totalDataRows(), report.getSuccessRows());
            assertEquals(FixtureCatalog.totalDataRows(), report.getTotalRows());
            assertEquals(0, report.getErrorRows());
            assertNull(report.getData());
            assertTrue(report.getGlobalErrors().isEmpty());
            assertTrue(options.isFailFastOnTemplateMismatch());
            assertEquals(100, options.getMaxErrorRows());
            ImportReportAssertions.assertBalancedStatistics(report);
            log.info("[DONE] API preset report passed");
        }
    }

    /** 分支：模板、阈值、异常包装、数据收集上限等。 */
    @Nested
    class Branch {

        /** 模板列缺失 + failFast：立即停止，不扫数据行。 */
        @Test
        void importReport_templateMismatchFailFast_stopsImmediately() throws Exception {
            log.info("[STEP] configure missing title with failFast=true");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportOptions options = new ImportOptions();
            options.setFailFastOnTemplateMismatch(true);

            ImportReport<Student> report = excelImport.switchSheet(0, Student::new)
                    .addConvert("missing-title-column", Student::setUserName)
                    .readWithReport(options);
            ImportReportAssertions.logReportSummary(report);

            assertTrue(report.getGlobalErrors().contains(ImportGlobalErrorEnum.TEMPLATE_MISMATCH));
            assertEquals(0, report.getSuccessRows());
            assertEquals(0, report.getTotalRows());
            log.info("[DONE] template mismatch fail-fast passed");
        }

        /** 每行后置抛错，达 maxErrorRows=5 后提前终止并写入 ERROR_LIMIT_REACHED。 */
        @Test
        void importReport_errorLimitReached_stopsBeforeFullScan() throws Exception {
            log.info("[STEP] inject CheckedExcelException on every row with maxErrorRows=5");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportOptions options = new ImportOptions();
            options.setMaxErrorRows(5);
            options.setCollectRowErrors(true);

            ImportReport<Student> report = ImportTestHelper.bindStandardColumns(
                            excelImport.switchSheet(0, Student::new))
                    .addConvertPostProcessor(student -> {
                        throw new CheckedExcelException("forced row validation error");
                    })
                    .readWithReport(options);
            ImportReportAssertions.logReportSummary(report);

            assertEquals(5, report.getErrorRows());
            assertTrue(report.getGlobalErrors().contains(ImportGlobalErrorEnum.ERROR_LIMIT_REACHED));
            assertTrue(report.getTotalRows() < FixtureCatalog.totalDataRows());
            assertEquals(5, report.getRowErrors().size());
            assertTrue(report.getRowErrors().get(0).getMessage().contains("[失败!]"));
            log.info("[DONE] error limit reached passed, scannedRows={}", report.getTotalRows());
        }

        /** 模板不匹配时行错误不入 rowErrors，但可同时出现模板与阈值全局错误。 */
        @Test
        void importReport_templateMismatchAndErrorLimit_bothPresentInGlobalErrors() throws Exception {
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportOptions options = new ImportOptions();
            options.setFailFastOnTemplateMismatch(false);
            options.setMaxErrorRows(5);
            options.setCollectRowErrors(true);

            ImportReport<Student> report = excelImport.switchSheet(0, Student::new)
                    .addConvert("missing-title-column", Student::setUserName)
                    .readWithReport(options);
            ImportReportAssertions.logReportSummary(report);

            assertTrue(report.getGlobalErrors().contains(ImportGlobalErrorEnum.TEMPLATE_MISMATCH));
            assertTrue(report.getGlobalErrors().contains(ImportGlobalErrorEnum.ERROR_LIMIT_REACHED));
            assertTrue(report.getRowErrors().isEmpty());
            log.info("[DONE] combined global errors passed");
        }

        /** collectSuccessData + maxSuccessData：成功条数全量统计，对象列表最多 100 条。 */
        @Test
        void importReport_collectSuccessData_respectsMaxSuccessDataCap() throws Exception {
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportOptions options = new ImportOptions();
            options.setCollectSuccessData(true);
            options.setMaxSuccessData(100);

            ImportReport<Student> report = ImportTestHelper.bindStandardColumns(
                            excelImport.switchSheet(0, Student::new))
                    .readWithReport(options);
            ImportReportAssertions.logReportSummary(report);

            assertNotNull(report.getData());
            assertEquals(100, report.getData().size());
            assertEquals(FixtureCatalog.totalDataRows(), report.getSuccessRows());
            assertNotNull(report.getData().get(0).getUserName());
            assertNotNull(report.getData().get(99).getEmail());
            log.info("[DONE] success data cap passed, collected={}", report.getData().size());
        }

        /** 已注册的 TestBizException：记错行但不中断，后续行继续读。 */
        @Test
        void importReport_customBizException_doesNotInterruptRemainingRows() throws Exception {
            AtomicInteger visited = new AtomicInteger();
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportReport<Student> report = ImportTestHelper.bindStandardColumns(
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
            ImportReport<Student> report = ImportTestHelper.bindStandardColumns(
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

        /** 年龄>90 抛 CheckedExcelException：total = success + error，行错误条数一致。 */
        @Test
        void importReport_statistics_totalEqualsSuccessPlusError() throws Exception {
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportReport<Student> report = ImportTestHelper.bindStandardColumns(
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

        /** forExcelResponse 预设：读全表、不 fail-fast，成功路径下 rowErrors 列表可初始化。 */
        @Test
        void importReport_excelResponsePreset_collectsRowErrorsOnSuccessPath() throws Exception {
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportReport<Student> report = ImportTestHelper.bindStandardColumns(
                            excelImport.switchSheet(0, Student::new))
                    .readWithReport(ImportOptions.forExcelResponse());
            ImportReportAssertions.logReportSummary(report);

            assertEquals(FixtureCatalog.totalDataRows(), report.getSuccessRows());
            assertNotNull(report.getRowErrors());
            assertFalse(ImportOptions.forExcelResponse().isFailFastOnTemplateMismatch());
            ImportReportAssertions.assertNoErrors(report);
            log.info("[DONE] forExcelResponse preset passed");
        }
    }
}
