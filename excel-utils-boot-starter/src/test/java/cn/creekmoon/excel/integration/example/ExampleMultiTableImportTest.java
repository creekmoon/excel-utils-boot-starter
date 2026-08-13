package cn.creekmoon.excel.integration.example;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.Teacher;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

import java.io.InputStream;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

/**
 * 复刻 excel-utils-example /import/with-reset：同一 Sheet 中读取三个连续表格。
 *
 * <p>fixture（导入测试Sheet页存在多个表格.xlsx）实际布局：
 * <ul>
 *   <li>row 0-50：Student 表（表头 row0，数据 1..50 共 50 行）</li>
 *   <li>row 51：空行分隔</li>
 *   <li>row 52-82：Teacher 表（表头 row52，数据 53..82 共 30 行）</li>
 *   <li>row 83：空行分隔</li>
 *   <li>row 85-95：Student 多级表头表（表头 row85 为「基本信息/年龄」，数据 86..95 共 10 行）</li>
 * </ul>
 * 注：示例代码中的 range(84) 与 fixture 实际表头行 85 错位，此处按 fixture 实际布局复刻。
 */
class ExampleMultiTableImportTest {

    private static final String FIXTURE = "导入测试Sheet页存在多个表格.xlsx";

    @Test
    void reset_readsThreeTablesInSameSheet() throws Exception {
        ExcelImport excelImport = ImportTestHelper.createImport(FIXTURE);
        AtomicReference<Student> firstStudent = new AtomicReference<>();
        AtomicReference<Student> lastStudentOfTable3 = new AtomicReference<>();

        /* 表格 1：Student（row 0 表头，数据 1..50） */
        TitleReader<Student> reader = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .range(0, 50);
        reader.read(s -> {
            if (firstStudent.get() == null) {
                firstStudent.set(s);
            }
        });
        ImportReport studentReport = reader.getReport();
        ImportReportAssertions.logReportSummary(studentReport);
        assertEquals(50, studentReport.getTotalRows(), "表格 1 应读到 50 行");
        ImportReportAssertions.assertNoErrors(studentReport);

        /* 表格 2：Teacher（row 52 表头，数据 53..82 共 30 行） */
        TitleReader<Teacher> teacherReader = reader.reset(Teacher::new)
                .addConvert("教师姓名", Teacher::setTeacherName)
                .addConvert("工作年限", IntegerConverter::parse, Teacher::setWorkYear)
                .range(52, 82);
        teacherReader.read(x -> {});
        ImportReport teacherReport = teacherReader.getReport();
        ImportReportAssertions.logReportSummary(teacherReport);
        assertEquals(30, teacherReport.getTotalRows(), "表格 2 应读到 30 行");
        ImportReportAssertions.assertNoErrors(teacherReport);

        /* 表格 3：Student 多级表头（row 85 表头，数据 86..95 共 10 行） */
        TitleReader<Student> thirdReader = teacherReader.reset(Student::new)
                .addConvert("基本信息", Student::setUserName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .range(85);
        thirdReader.read(s -> {
            if (lastStudentOfTable3.get() == null) {
                lastStudentOfTable3.set(s);
            }
        });
        ImportReport thirdReport = thirdReader.getReport();
        ImportReportAssertions.logReportSummary(thirdReport);
        assertEquals(10, thirdReport.getTotalRows(), "表格 3 应读到 10 行");
        ImportReportAssertions.assertNoErrors(thirdReport);

        /* 抽样断言：导入值与源文件 POI 预读一致 */
        try (InputStream in = ImportTestHelper.class.getResourceAsStream("/fixtures/import/" + FIXTURE);
             Workbook workbook = WorkbookFactory.create(in)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row table1First = sheet.getRow(1);
            assertEquals(table1First.getCell(0).getStringCellValue(), firstStudent.get().getUserName(),
                    "表格 1 首行用户名应与源文件一致");
            assertEquals(Double.valueOf(table1First.getCell(2).getNumericCellValue()).intValue(),
                    firstStudent.get().getAge(), "表格 1 首行年龄应与源文件一致");

            Row table3First = sheet.getRow(86);
            assertEquals(table3First.getCell(0).getStringCellValue(), lastStudentOfTable3.get().getUserName(),
                    "表格 3 首行用户名应与源文件一致");
            assertNotNull(lastStudentOfTable3.get().getAge());
        }
    }
}
