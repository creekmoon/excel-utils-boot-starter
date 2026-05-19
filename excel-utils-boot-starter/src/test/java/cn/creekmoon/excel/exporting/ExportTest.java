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

/** Excel 导出：多 sheet、多级表头、条件样式、reset 分段写。 */
@Slf4j
class ExportTest {

    /** 「基本信息::列名」多级表头占用的行数（含分组行 + 子标题行）。 */
    private static final int MULTI_HEADER_DEPTH = 2;

    @Nested
    class Smoke {

        /** 6 万 + 1 万行双 sheet：行数、分组表头、首行抽样单元格。 */
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

        /** 两 sheet 交替 write，各自累计行数（各 2000 数据行 + 表头）。 */
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

        /** 先写空列表（仅表头），再分批追加，数据从表头下一行起连续排列。 */
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

        /** 多级表头：存在合并区，分组与子列标题文本正确。 */
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

        /** 条件样式：全名列匹配时填充 LIGHT_ORANGE 前景色。 */
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

        /** 同一 sheet 上 reset 切换 Student/Teacher 表结构，多段表格首尾相接。 */
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

        /** range 严格截断：写入 100 行但限制 lastDataRowIndex=5，只应保留 5 行数据。 */
        @Test
        void export_rangeLimit_truncatesExcessRows() throws Exception {
            ArrayList<Student> batch = StudentFixtures.createList(100);
            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet(Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("年龄", Student::getAge)
                    .range(0, 1, 5)
                    .write(batch);

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                // 标题 1 行 + 数据最多 5 行 = 6 行
                PoiWorkbookAssertions.assertRowCount(sheet, 6);
                PoiWorkbookAssertions.assertCellString(sheet, 1, 0, batch.get(0).getUserName());
                PoiWorkbookAssertions.assertCellString(sheet, 5, 0, batch.get(4).getUserName());
                // 第 6 行及以后不应存在有效数据
                PoiWorkbookAssertions.assertCellBlank(sheet, 6, 0);
            }
            log.info("[DONE] range limit truncates excess rows passed");
        }

        /** 同一列多个互斥条件样式：年龄>50 橙色，年龄<30 绿色，中间年龄段无样式。 */
        @Test
        void export_multipleExclusiveConditionalStyles_appliesOnlyMatched() throws Exception {
            ArrayList<Student> batch = StudentFixtures.createList(100);
            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet(Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("年龄", Student::getAge,
                            of(LIGHT_ORANGE, s -> s.getAge() > 50),
                            of(LIGHT_GREEN, s -> s.getAge() < 30))
                    .write(batch);

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                for (int row = 1; row <= 100; row++) {
                    Student s = batch.get(row - 1);
                    if (s.getAge() > 50) {
                        PoiWorkbookAssertions.assertCellFillForeground(sheet, row, 1, IndexedColors.LIGHT_ORANGE);
                    } else if (s.getAge() < 30) {
                        PoiWorkbookAssertions.assertCellFillForeground(sheet, row, 1, IndexedColors.LIGHT_GREEN);
                    }
                    // 30 <= age <= 50 不断言颜色（保持默认）
                }
                log.info("[ASSERT] multiple exclusive conditional styles verified on 100 rows");
            }
            log.info("[DONE] multiple exclusive conditional styles passed");
        }

        /** 空 sheet 穿插：Sheet0 有数据+条件样式，Sheet1 仅表头，Sheet2 有数据+不同条件样式。 */
        @Test
        void export_interleavedEmptySheet_conditionalStyleIsolated() throws Exception {
            ArrayList<Student> students = StudentFixtures.createList(10);
            ArrayList<Teacher> teachers = TeacherFixtures.createList(10);

            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet("Students", Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("年龄", Student::getAge, of(LIGHT_ORANGE, s -> s.getAge() > 25))
                    .write(students);

            excelExport.switchNewSheet("Empty", Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("年龄", Student::getAge, of(LIGHT_GREEN, s -> true))
                    .write(Collections.emptyList());

            excelExport.switchNewSheet("Teachers", Teacher.class)
                    .addTitle("教师姓名", Teacher::getTeacherName)
                    .addTitle("工作年限", Teacher::getWorkYear, of(LIGHT_GREEN, t -> t.getWorkYear() > 5))
                    .write(teachers);

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                PoiWorkbookAssertions.assertSheetNames(workbook, "Students", "Empty", "Teachers");

                Sheet sheet0 = workbook.getSheetAt(0);
                Sheet sheet1 = workbook.getSheetAt(1);
                Sheet sheet2 = workbook.getSheetAt(2);

                // Sheet0: 10 行数据，条件样式生效（age>25 的是索引 6-9，对应 Excel 行 7-10）
                PoiWorkbookAssertions.assertRowCount(sheet0, 11);
                PoiWorkbookAssertions.assertCellFillForeground(sheet0, 7, 1, IndexedColors.LIGHT_ORANGE);
                // age<=25 的行保持默认色（不断言）

                // Sheet1: 仅表头 1 行，无数据，不应 crash
                PoiWorkbookAssertions.assertRowCount(sheet1, 1);

                // Sheet2: 10 行数据，条件样式生效且不污染 Sheet0
                // Teacher workYear = index（0-9），>5 的是索引 6-9，对应 Excel 行 7-10
                PoiWorkbookAssertions.assertRowCount(sheet2, 11);
                PoiWorkbookAssertions.assertCellFillForeground(sheet2, 7, 1, IndexedColors.LIGHT_GREEN);
            }
            log.info("[DONE] interleaved empty sheet with isolated conditional styles passed");
        }

        /**
         * reset + 单级表头续写：第一段单级表头+数据，reset 后第二段单级表头+数据。
         * 注：EasyExcel 对同一 sheet 复用 WriteSheetHolder，reset 后多级表头无法重新写入，
         *     故本测试使用单级表头验证数据接续行为。
         */
        @Test
        void export_resetSingleLevelHeader_dataContinuesAtCorrectPosition() throws Exception {
            ArrayList<Student> students = StudentFixtures.createList(20);

            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet("Continued", Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("年龄", Student::getAge)
                    .write(students.subList(0, 10))
                    .reset(Student.class)
                    .addTitle("邮箱", Student::getEmail)
                    .addTitle("生日", Student::getBirthday)
                    .write(students.subList(10, 20));

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                // 第一段：1 行表头 + 10 行数据 = 11 行
                // 第二段：数据从 firstRowIndex 开始接续（单级表头深度=1，从 row 11 开始）
                // 总行数 = 11 + 10 = 21
                PoiWorkbookAssertions.assertRowCount(sheet, 21);

                // 第一段表头在第 0 行
                PoiWorkbookAssertions.assertCellString(sheet, 0, 0, "用户名");
                // 第一段数据从第 1 行开始
                PoiWorkbookAssertions.assertCellString(sheet, 1, 0, students.get(0).getUserName());

                // 第二段数据从第 11 行开始接续
                PoiWorkbookAssertions.assertCellString(sheet, 11, 0, students.get(10).getEmail());
                PoiWorkbookAssertions.assertCellString(sheet, 20, 0, students.get(19).getEmail());

                log.info("[ASSERT] reset single level header: totalRows={}, firstSegmentEnd={}, secondSegmentStart={}",
                        sheet.getLastRowNum() + 1, 10, 11);
            }
            log.info("[DONE] reset single level header data continuation passed");
        }

        /** 大数据量 + 全列条件样式：1 万行 × 4 列样式，验证性能与内存。 */
        @Test
        void export_largeDataWithConditionalStyles_allRowsStyledCorrectly() throws Exception {
            int rows = 10_000;
            ArrayList<Student> batch = StudentFixtures.createList(rows);
            log.info("[STEP] export {} rows with conditional styles on 4 columns", rows);

            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet(Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("全名", Student::getFullName, of(LIGHT_ORANGE, s -> s.getUserName().endsWith("0")))
                    .addTitle("年龄", Student::getAge, of(LIGHT_GREEN, s -> s.getAge() > 50))
                    .addTitle("邮箱", Student::getEmail, of(LIGHT_ORANGE, s -> s.getEmail().contains("99")))
                    .write(batch);

            File file = excelExport.stopWrite();
            assertTrue(file.exists());
            log.info("[RESULT] export file size={} bytes", file.length());

            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                PoiWorkbookAssertions.assertRowCount(sheet, rows + 1);

                // 抽样验证：第 1、1000、5000、10000 行
                int[] sampleRows = {1, 1000, 5000, 10000};
                for (int row : sampleRows) {
                    Student s = batch.get(row - 1);
                    if (s.getUserName().endsWith("0")) {
                        PoiWorkbookAssertions.assertCellFillForeground(sheet, row, 1, IndexedColors.LIGHT_ORANGE);
                    }
                    if (s.getAge() > 50) {
                        PoiWorkbookAssertions.assertCellFillForeground(sheet, row, 2, IndexedColors.LIGHT_GREEN);
                    }
                    if (s.getEmail().contains("99")) {
                        PoiWorkbookAssertions.assertCellFillForeground(sheet, row, 3, IndexedColors.LIGHT_ORANGE);
                    }
                }
                log.info("[ASSERT] large data conditional styles sampled at rows {}", (Object) sampleRows);
            }
            log.info("[DONE] large data with conditional styles passed");
        }
    }
}
