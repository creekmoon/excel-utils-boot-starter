package cn.creekmoon.excel.integration.example;

import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle;
import cn.creekmoon.excel.support.assertion.PoiWorkbookAssertions;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.StudentFixtures;
import cn.creekmoon.excel.support.fixture.Teacher;
import cn.creekmoon.excel.support.fixture.TeacherFixtures;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.ArrayList;

import static cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle.of;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.LIGHT_GREEN;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.LIGHT_ORANGE;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 复刻 excel-utils-example 导出案例：export/basic（多 sheet + 中文 sheet 名 + 多级表头 + 日期列）、
 * export/with-style（自定义数字格式 + 条件样式）、export/with-range（三种 range 组合）。
 */
class ExampleExportTest {

    @Nested
    class Basic {

        /** 复刻 /export/basic：10k 学生 + 5k 教师，双 sheet 中文名，多级表头含日期列。 */
        @Test
        void exportBasic_twoSheetsChineseNames_multiLevelHeaderWithDates() throws Exception {
            ArrayList<Student> students = StudentFixtures.createList(10_000);
            ArrayList<Teacher> teachers = TeacherFixtures.createList(5_000);

            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet("学生信息", Student.class)
                    .addTitle("基本信息::用户名", Student::getUserName)
                    .addTitle("基本信息::全名", Student::getFullName)
                    .addTitle("联系方式::邮箱", Student::getEmail)
                    .addTitle("其他::年龄", Student::getAge)
                    .addTitle("其他::生日", Student::getBirthday)
                    .addTitle("其他::过期时间", Student::getExpTime)
                    .write(students);
            excelExport.switchNewSheet("教师信息", Teacher.class)
                    .addTitle("教师姓名", Teacher::getTeacherName)
                    .addTitle("工作年限", Teacher::getWorkYear)
                    .write(teachers);

            File file = excelExport.stopWrite();
            assertTrue(file.exists());
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                PoiWorkbookAssertions.assertSheetNames(workbook, "学生信息", "教师信息");

                /* Sheet0：多级表头 2 行 + 10000 数据 */
                Sheet sheet0 = workbook.getSheetAt(0);
                PoiWorkbookAssertions.assertRowCount(sheet0, 10_000 + 2);
                PoiWorkbookAssertions.assertCellString(sheet0, 0, 0, "基本信息");
                PoiWorkbookAssertions.assertCellString(sheet0, 1, 0, "用户名");
                PoiWorkbookAssertions.assertCellString(sheet0, 2, 0, students.get(0).getUserName());
                PoiWorkbookAssertions.assertCellString(sheet0, 2, 2, students.get(0).getEmail());
                /* 抽样：首行/中间/末行数据 */
                PoiWorkbookAssertions.assertCellString(sheet0, 2, 3, String.valueOf(students.get(0).getAge()));
                PoiWorkbookAssertions.assertCellString(sheet0, 5001, 3, String.valueOf(students.get(4999).getAge()));
                PoiWorkbookAssertions.assertCellString(sheet0, 10001, 3, String.valueOf(students.get(9999).getAge()));
                /* 日期列导出为日期单元格 */
                Cell birthdayCell = sheet0.getRow(2).getCell(4);
                assertEquals(org.apache.poi.ss.usermodel.CellType.NUMERIC, birthdayCell.getCellType(),
                        "生日列应为日期单元格");
                assertTrue(org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(birthdayCell));

                /* Sheet1：1 行表头 + 5000 数据 */
                Sheet sheet1 = workbook.getSheetAt(1);
                PoiWorkbookAssertions.assertRowCount(sheet1, 5_000 + 1);
                PoiWorkbookAssertions.assertCellString(sheet1, 1, 0, teachers.get(0).getTeacherName());
                PoiWorkbookAssertions.assertCellString(sheet1, 5000, 0, teachers.get(4999).getTeacherName());
            }
        }
    }

    @Nested
    class WithStyle {

        /** 复刻 /export/with-style：自定义 #,##0.00 格式 + LIGHT_GREEN 条件着色。 */
        @Test
        void exportWithStyle_customFormatAndConditionalColor() throws Exception {
            ArrayList<Student> students = StudentFixtures.createList(1_000);
            ExcelCellStyle twoDecimalStyle = new ExcelCellStyle(
                    (workbook, style) -> style.setDataFormat(
                            workbook.createDataFormat().getFormat("#,##0.00")));

            ExcelExport excelExport = ExcelExport.create();
            excelExport.switchNewSheet(Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("全名", Student::getFullName, of(LIGHT_GREEN, s -> s.getAge() > 50))
                    .addTitle("年龄", Student::getAge, of(twoDecimalStyle, s -> s.getAge() > 50))
                    .addTitle("邮箱", Student::getEmail)
                    .write(students);

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                Sheet sheet = workbook.getSheetAt(0);
                int styledRows = 0;
                for (int row = 1; row <= 1000; row++) {
                    Student s = students.get(row - 1);
                    if (s.getAge() > 50) {
                        styledRows++;
                        PoiWorkbookAssertions.assertCellFillForeground(sheet, row, 1, IndexedColors.LIGHT_GREEN);
                        assertTrue(sheet.getRow(row).getCell(2).getCellStyle().getDataFormatString().contains("#,##0.00"),
                                "年龄列应应用 #,##0.00 格式");
                    }
                }
                assertTrue(styledRows > 0, "年龄>50 的行应存在");
            }
        }
    }

    @Nested
    class WithRange {

        /** 复刻 /export/with-range 三种 range 用法在同一文件中的落盘位置。 */
        @Test
        void exportWithRange_threeRangeVariants_placedAtExpectedPositions() throws Exception {
            ArrayList<Student> students = StudentFixtures.createList(100);

            ExcelExport excelExport = ExcelExport.create();
            /* 示例 1: range(5, 19) — 表头 row5，数据 6..19 共 14 行 */
            excelExport.switchNewSheet("范围示例1", Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("全名", Student::getFullName)
                    .addTitle("年龄", Student::getAge)
                    .range(5, 19)
                    .write(students);
            /* 示例 2: range(9) — 表头 row9，数据从 10 行开始不限制 */
            excelExport.switchNewSheet("范围示例2", Student.class)
                    .addTitle("用户名", Student::getUserName)
                    .addTitle("全名", Student::getFullName)
                    .addTitle("年龄", Student::getAge, of(LIGHT_GREEN, s -> s.getAge() > 50))
                    .range(9)
                    .write(students);
            /* 示例 3: range(2, 4, 15) 多级表头 — 表头 row2-3，数据 4..15 共 12 行 */
            excelExport.switchNewSheet("多级表头+范围", Student.class)
                    .addTitle("基本信息::用户名", Student::getUserName)
                    .addTitle("基本信息::全名", Student::getFullName, of(LIGHT_ORANGE, s -> true))
                    .addTitle("年龄", Student::getAge)
                    .range(2, 4, 15)
                    .write(students);

            File file = excelExport.stopWrite();
            try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
                PoiWorkbookAssertions.assertSheetNames(workbook, "范围示例1", "范围示例2", "多级表头+范围");

                /* 示例 1：row0-4 空，表头 row5，数据 6..19，共 20 行 */
                Sheet sheet0 = workbook.getSheetAt(0);
                PoiWorkbookAssertions.assertRowCount(sheet0, 20);
                PoiWorkbookAssertions.assertCellString(sheet0, 5, 0, "用户名");
                PoiWorkbookAssertions.assertCellString(sheet0, 6, 0, students.get(0).getUserName());
                PoiWorkbookAssertions.assertCellString(sheet0, 19, 0, students.get(13).getUserName());

                /* 示例 2：表头 row9，数据 10..109，共 110 行 */
                Sheet sheet1 = workbook.getSheetAt(1);
                PoiWorkbookAssertions.assertRowCount(sheet1, 110);
                PoiWorkbookAssertions.assertCellString(sheet1, 9, 0, "用户名");
                PoiWorkbookAssertions.assertCellString(sheet1, 10, 0, students.get(0).getUserName());
                PoiWorkbookAssertions.assertCellString(sheet1, 109, 0, students.get(99).getUserName());

                /* 示例 3：多级表头 row2-3，数据 4..15，共 16 行 */
                Sheet sheet2 = workbook.getSheetAt(2);
                PoiWorkbookAssertions.assertRowCount(sheet2, 16);
                PoiWorkbookAssertions.assertCellString(sheet2, 2, 0, "基本信息");
                PoiWorkbookAssertions.assertCellString(sheet2, 3, 0, "用户名");
                PoiWorkbookAssertions.assertCellString(sheet2, 4, 0, students.get(0).getUserName());
                PoiWorkbookAssertions.assertCellString(sheet2, 15, 0, students.get(11).getUserName());
                PoiWorkbookAssertions.assertCellFillForeground(sheet2, 4, 1, IndexedColors.LIGHT_ORANGE);
            }
        }
    }
}
