package cn.creekmoon.excel.exporting;

import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.core.W.title.TitleWriter;
import cn.creekmoon.excel.support.assertion.PoiWorkbookAssertions;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.StudentFixtures;
import cn.creekmoon.excel.support.fixture.Teacher;
import cn.creekmoon.excel.support.fixture.TeacherFixtures;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;

import static cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle.of;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.LIGHT_GREEN;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.LIGHT_ORANGE;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

@Slf4j
class ExportTest {

    private static final int MULTI_HEADER_DEPTH = 2;

    @Nested
    class Smoke {

        @Test
        void export_multiSheetMultiLevelHeader_largeData_hasCorrectStructure() throws Exception {
            int sheet0Rows = 60_000;
            int sheet1Rows = 10_000;
            ArrayList<Student> sheet0Data = StudentFixtures.createList(sheet0Rows);
            ArrayList<Student> sheet1Data = StudentFixtures.createList(sheet1Rows);
            log.info("[STEP] export sheet0Rows={}, sheet1Rows={}", sheet0Rows, sheet1Rows);

            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet(Student.class)
                    .addTitle("基本信息::用户名", Student::getUserName)
                    .addTitle("基本信息::全名", Student::getFullName, of(LIGHT_ORANGE, x -> true))
                    .addTitle("年龄", Student::getAge)
                    .addTitle("邮箱", Student::getEmail)
                    .write(sheet0Data);

            excelExport.switchNewSheet(Student.class)
                    .addTitle("年龄", Student::getAge)
                    .addTitle("邮箱", Student::getEmail)
                    .write(sheet1Data);

            File file = excelExport.stopWrite();
            log.info("[RESULT] export file={}, size={} bytes", file.getAbsolutePath(), file.length());
            assertTrue(file.exists());
            assertTrue(file.length() > 0);

            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                PoiWorkbookAssertions.assertSheetNames(workbook, "Sheet0", "Sheet1");
                Sheet sheet0 = workbook.getSheetAt(0);
                Sheet sheet1 = workbook.getSheetAt(1);
                PoiWorkbookAssertions.assertRowCount(sheet0, sheet0Rows + MULTI_HEADER_DEPTH);
                PoiWorkbookAssertions.assertRowCount(sheet1, sheet1Rows + 1);
                PoiWorkbookAssertions.assertCellString(sheet0, 0, 0, "基本信息");
                PoiWorkbookAssertions.assertCellString(sheet0, MULTI_HEADER_DEPTH, 0, sheet0Data.get(0).getUserName());
                PoiWorkbookAssertions.assertCellString(sheet0, MULTI_HEADER_DEPTH, 2, String.valueOf(sheet0Data.get(0).getAge()));
                PoiWorkbookAssertions.assertCellString(sheet1, 1, 0, String.valueOf(sheet1Data.get(0).getAge()));
                log.info("[ASSERT] sheet0 lastRow={}, sheet1 lastRow={}", sheet0.getLastRowNum(), sheet1.getLastRowNum());
            }
            log.info("[DONE] large multi-sheet export passed");
        }
    }

    @Nested
    class Branch {

        @Test
        void export_interleavedMultiSheetWrite_accumulatesRowCounts() throws Exception {
            ArrayList<Student> batch = StudentFixtures.createList(1_000);
            ExcelExport excelExport = ExcelExport.create();
            TitleWriter<Student> writer0 = excelExport.switchNewSheet(Student.class)
                    .addTitle("基本信息::用户名", Student::getUserName)
                    .addTitle("基本信息::全名", Student::getFullName, of(LIGHT_ORANGE, x -> true))
                    .addTitle("年龄", Student::getAge)
                    .addTitle("邮箱", Student::getEmail);
            TitleWriter<Student> writer1 = excelExport.switchNewSheet(Student.class)
                    .addTitle("年龄", Student::getAge)
                    .addTitle("邮箱", Student::getEmail);

            log.info("[STEP] interleaved write 1000 rows twice on two sheets");
            writer0.write(batch);
            writer1.write(batch);
            writer0.write(batch);
            writer1.write(batch);

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                PoiWorkbookAssertions.assertRowCount(workbook.getSheetAt(0), 1_000 * 2 + MULTI_HEADER_DEPTH);
                PoiWorkbookAssertions.assertRowCount(workbook.getSheetAt(1), 1_000 * 2 + 1);
            }
            log.info("[DONE] interleaved export passed");
        }

        @Test
        void export_emptyWriteThenAppend_startsDataAfterHeader() throws Exception {
            ArrayList<Student> batch = StudentFixtures.createList(10);
            ExcelExport excelExport = ExcelExport.create();
            TitleWriter<Student> writer = excelExport.switchNewSheet(Student.class)
                    .addTitle("基本信息::用户名", Student::getUserName)
                    .addTitle("基本信息::全名", Student::getFullName, of(LIGHT_ORANGE, x -> true))
                    .addTitle("年龄", Student::getAge);

            log.info("[STEP] write empty header only, then append data twice");
            writer.write(Collections.emptyList());
            writer.write(batch);
            writer.write(batch);

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                PoiWorkbookAssertions.assertRowCount(sheet, 20 + MULTI_HEADER_DEPTH);
                PoiWorkbookAssertions.assertCellString(sheet, MULTI_HEADER_DEPTH, 0, batch.get(0).getUserName());
                PoiWorkbookAssertions.assertCellString(sheet, MULTI_HEADER_DEPTH + 10, 0, batch.get(0).getUserName());
            }
            log.info("[DONE] empty-then-append export passed");
        }

        @Test
        void export_multiLevelHeader_hasMergedRegionsAndHeaderText() throws Exception {
            ArrayList<Student> batch = StudentFixtures.createList(1);
            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet(Student.class)
                    .addTitle("基本信息::用户名", Student::getUserName)
                    .addTitle("基本信息::全名", Student::getFullName)
                    .addTitle("年龄", Student::getAge)
                    .write(batch);

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                assertTrue(sheet.getNumMergedRegions() > 0);
                PoiWorkbookAssertions.assertCellString(sheet, 0, 0, "基本信息");
                PoiWorkbookAssertions.assertCellString(sheet, 1, 0, "用户名");
                PoiWorkbookAssertions.assertCellString(sheet, 1, 1, "全名");
            }
            log.info("[DONE] multi-level header export passed, mergedRegions verified");
        }

        @Test
        void export_conditionalStyle_appliesForegroundColorWhenMatched() throws Exception {
            ArrayList<Student> batch = StudentFixtures.createList(3);
            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet(Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("全名", Student::getFullName, of(LIGHT_ORANGE, x -> true))
                    .write(batch);

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                for (int row = 1; row <= 3; row++) {
                    PoiWorkbookAssertions.assertCellFillForeground(sheet, row, 1, IndexedColors.LIGHT_ORANGE);
                }
                log.info("[ASSERT] conditional style applied on rows 1-3 col 1");
            }
            log.info("[DONE] conditional style export passed");
        }

        @Test
        void export_resetSameSheetMultipleSegments_writesAllTables() throws Exception {
            ArrayList<Student> students = StudentFixtures.createList(50);
            ArrayList<Teacher> teachers = TeacherFixtures.createList(30);

            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet("MultiTableSheet", Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("全名", Student::getFullName)
                    .addTitle("年龄", Student::getAge)
                    .write(students)
                    .reset(Teacher.class)
                    .addTitle("教师姓名", Teacher::getTeacherName)
                    .addTitle("工作年限", Teacher::getWorkYear, of(LIGHT_GREEN, x -> x.getWorkYear() > 5))
                    .write(teachers)
                    .reset(Student.class, 2)
                    .addTitle("基本信息::用户名", Student::getUserName)
                    .addTitle("基本信息::全名", Student::getFullName)
                    .addTitle("年龄", Student::getAge)
                    .write(students.subList(0, 10));

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                assertEquals("MultiTableSheet", workbook.getSheetName(0));
                assertTrue(sheet.getLastRowNum() + 1 > students.size() + teachers.size());
                PoiWorkbookAssertions.assertCellString(sheet, 0, 0, "用户名");
                log.info("[RESULT] reset same sheet totalRows={}", sheet.getLastRowNum() + 1);
            }
            log.info("[DONE] reset multi-segment export passed");
        }
    }
}
