package cn.creekmoon.excel.importing;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.assertion.PoiWorkbookAssertions;
import cn.creekmoon.excel.support.fixture.FixtureCatalog;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.ExcelFileUtils;
import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.hutool.core.io.FileUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

@Slf4j
class ImportTitleReaderTest {

    @Nested
    class Smoke {

        @Test
        void importTitleReader_standardTemplate_parsesAllRowsWithCorrectStatistics() throws Exception {
            log.info("[STEP] load fixture={}", FixtureCatalog.STANDARD_IMPORT);
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            TitleReader<Student> reader = ImportTestHelper.bindStandardColumns(
                    excelImport.switchSheet(0, Student::new));

            log.info("[STEP] read all rows from sheet0");
            List<Student> dataList = reader.getAll();
            ImportReport<Student> report = reader.getReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(FixtureCatalog.totalDataRows(), dataList.size(), "parsed row count");
            for (int excelRow : new int[]{6, 80, 195, 631}) {
                FixtureCatalog.SampleRow sample = FixtureCatalog.sampleRow(excelRow);
                Student actual = dataList.get(sample.excelRowNumber() - 2);
                log.info("[ASSERT] spot-check excelRow={}, userName={}, age={}", excelRow, actual.getUserName(), actual.getAge());
                assertEquals(sample.userName(), actual.getUserName());
                assertEquals(sample.fullName(), actual.getFullName());
                assertEquals(sample.email(), actual.getEmail());
                assertEquals(sample.age(), actual.getAge());
            }

            assertEquals(1, report.getProcessedFirstRowIndex());
            assertEquals(FixtureCatalog.totalDataRows(), report.getProcessedLastRowIndex());
            ImportReportAssertions.assertBalancedStatistics(report);
            ImportReportAssertions.assertNoErrors(report);
            assertEquals(FixtureCatalog.ageGreaterThan60Count(),
                    dataList.stream().filter(x -> x.getAge() > 60).count());
            assertTrue(dataList.stream().allMatch(s -> s.getUserName() != null && s.getEmail() != null));
            log.info("[DONE] standard import passed, rows={}", dataList.size());
        }

        @Test
        void importTitleReader_businessValidationFailure_writesResultColumnWithReason() throws Exception {
            log.info("[STEP] load fixture and import sheet0 as baseline");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            ImportTestHelper.bindStandardColumns(excelImport.switchSheet(0, Student::new)).getAll();

            log.info("[STEP] import sheet1 and inject business error on lisi@qq.com");
            TitleReader<Student> sheet2 = excelImport.switchSheet(1, Student::new)
                    .addConvert("生日", DateConverter::parse, Student::setBirthday)
                    .addConvert("邮箱", Student::setEmail);
            sheet2.read(student -> {
                if ("lisi@qq.com".equals(student.getEmail())) {
                    throw new CheckedExcelException("validation failed on lisi@qq.com");
                }
            });

            ImportReport<Student> report = sheet2.getReport();
            ImportReportAssertions.logReportSummary(report);
            assertEquals(1, report.getErrorRows());
            assertNotNull(report.getRowErrors());
            assertEquals(1, report.getRowErrors().size());
            assertTrue(report.getRowErrors().get(0).getMessage().contains("[失败!]"));

            log.info("[STEP] generate result file");
            File file = excelImport.generateResultFile(false);
            assertTrue(FileUtil.exist(file));
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet errorSheet = workbook.getSheetAt(1);
                Map<Integer, String> resultMessages = PoiWorkbookAssertions.readResultColumn(
                        errorSheet, 0, ExcelConstants.RESULT_TITLE);
                log.info("[RESULT] result column messages on sheet1: {}", resultMessages);
                assertEquals(1, resultMessages.size());
                assertTrue(resultMessages.values().iterator().next().contains("[失败!]"));

                Sheet sheet0 = workbook.getSheetAt(0);
                int resultColOnSheet0 = PoiWorkbookAssertions.findHeaderColumnIndex(sheet0, 0, ExcelConstants.RESULT_TITLE);
                PoiWorkbookAssertions.assertCellBlank(sheet0, 1, resultColOnSheet0);
                log.info("[ASSERT] success row on sheet0 keeps result column blank");
            }
            ExcelFileUtils.cleanTempFileByPathNow(file.getPath());
            log.info("[DONE] writeback validation passed");
        }
    }

    @Nested
    class Branch {

        @Test
        void importTitleReader_multiSheetSwitch_parsesEachSheetIndependently() throws Exception {
            log.info("[STEP] import sheet0 with callback counter");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            AtomicInteger sheet0Count = new AtomicInteger();
            TitleReader<Student> sheet0Reader = ImportTestHelper.bindStandardColumns(
                    excelImport.switchSheet(0, Student::new));
            sheet0Reader.read(x -> sheet0Count.incrementAndGet());
            assertEquals(FixtureCatalog.totalDataRows(), sheet0Count.get());

            log.info("[STEP] import sheet1");
            List<Student> sheet1Students = ImportTestHelper.bindStandardColumns(
                    excelImport.switchSheet(1, Student::new)).getAll();
            assertEquals(4, sheet1Students.size());
            assertTrue(sheet1Students.stream().allMatch(s -> s.getEmail() != null));

            ImportReportAssertions.logReportSummary(sheet0Reader.getReport());
            File file = excelImport.generateResultFile(false);
            assertTrue(file.length() > 0);
            ExcelFileUtils.cleanTempFileByPathNow(file.getPath());
            log.info("[DONE] multi-sheet import passed");
        }

        @Test
        void importTitleReader_readRangeLimited_returnsRowsInsideRangeOnly() throws Exception {
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            TitleReader<Student> reader = ImportTestHelper.bindStandardColumns(
                    excelImport.switchSheet(0, Student::new).range(0, 3, 4));

            log.info("[STEP] read with range titleRow=0, firstDataRow=3, lastDataRow=4");
            List<Student> students = reader.getAll();
            ImportReport<Student> report = reader.getReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(2, students.size());
            assertEquals(2, report.getTotalRows());
            assertEquals(2, report.getSuccessRows());
            log.info("[DONE] range-limited rows={}", students.size());
        }

        @Test
        void importTitleReader_getAllAndReport_keepsStatisticsConsistent() throws Exception {
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
            TitleReader<Student> reader = ImportTestHelper.bindStandardColumns(
                    excelImport.switchSheet(0, Student::new));

            List<Student> students = reader.getAll();
            ImportReport<Student> report = reader.getReport();
            ImportReportAssertions.logReportSummary(report);

            assertEquals(students.size(), report.getSuccessRows());
            assertEquals(FixtureCatalog.totalDataRows(), report.getTotalRows());
            assertEquals(0, report.getErrorRows());
            assertNull(report.getData());
            ImportReportAssertions.assertBalancedStatistics(report);
            log.info("[DONE] getAll/report consistency verified");
        }
    }
}
