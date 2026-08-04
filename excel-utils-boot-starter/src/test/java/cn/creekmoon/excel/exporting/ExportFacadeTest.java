package cn.creekmoon.excel.exporting;

import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.core.W.title.TitleWriter;
import cn.creekmoon.excel.support.assertion.PoiWorkbookAssertions;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.StudentFixtures;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockHttpServletResponse;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** ExcelExport 门面契约：sheet 下标连续性、复用、write 后锁定、下载头。 */
class ExportFacadeTest {

    @Test
    void switchSheet_withoutPriorSheets_throws() {
        ExcelExport excelExport = ExcelExport.create();
        // switchSheet(0) 前未调用 switchNewSheet：连续性检查失败
        assertThrows(RuntimeException.class, () -> excelExport.switchSheet(0, Student.class));
    }

    @Test
    void switchSheet_jumpingIndex_throws() {
        ExcelExport excelExport = ExcelExport.create();
        excelExport.switchNewSheet(Student.class).addTitle("年龄", Student::getAge)
                .write(Collections.emptyList());
        // 已存在 sheet0，直接切 sheet2（跳过 1）
        assertThrows(RuntimeException.class, () -> excelExport.switchSheet(2, Student.class));
    }

    @Test
    void switchSheet_existingIndex_returnsSameWriter() {
        ExcelExport excelExport = ExcelExport.create();
        TitleWriter<Student> first = excelExport.switchNewSheet(Student.class)
                .addTitle("年龄", Student::getAge);
        assertSame(first, excelExport.switchSheet(0, Student.class), "已存在 sheet 应复用同一 writer");
    }

    @Test
    void addTitle_afterWrite_throwsIllegalState() {
        ExcelExport excelExport = ExcelExport.create();
        TitleWriter<Student> writer = excelExport.switchNewSheet(Student.class)
                .addTitle("年龄", Student::getAge)
                .write(Collections.emptyList());
        assertThrows(IllegalStateException.class, () -> writer.addTitle("邮箱", Student::getEmail));
    }

    @Test
    void switchNewSheet_customName_keepsSheetName() throws Exception {
        ExcelExport excelExport = ExcelExport.create();
        excelExport.switchNewSheet("自定义表", Student.class)
                .addTitle("年龄", Student::getAge)
                .write(Collections.emptyList());

        File file = excelExport.stopWrite();
        try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
            assertEquals("自定义表", workbook.getSheetName(0));
        }
        assertTrue(file.exists());
    }

    @Test
    void response_setsContentDispositionWithExcelName() throws Exception {
        ExcelExport excelExport = ExcelExport.create();
        excelExport.switchNewSheet(Student.class)
                .addTitle("年龄", Student::getAge)
                .write(Collections.emptyList());

        MockHttpServletResponse response = new MockHttpServletResponse();
        excelExport.response(response);

        String disposition = response.getHeader("Content-Disposition");
        assertNotNull(disposition);
        assertTrue(disposition.contains(excelExport.excelName), disposition);
        assertTrue(excelExport.getTaskId() != null && !excelExport.getTaskId().isEmpty());
        assertTrue(excelExport.excelName.endsWith(".xlsx"));
    }

    @Test
    void stopWrite_withSingleSheetAndData_producesReadableFile() throws Exception {
        ArrayList<Student> data = StudentFixtures.createList(5);
        ExcelExport excelExport = ExcelExport.create();
        excelExport.switchNewSheet(Student.class)
                .addTitle("年龄", Student::getAge)
                .write(data);

        File file = excelExport.stopWrite();
        try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
            PoiWorkbookAssertions.assertRowCount(workbook.getSheetAt(0), 6);
            PoiWorkbookAssertions.assertCellString(workbook.getSheetAt(0), 1, 0, String.valueOf(data.get(0).getAge()));
        }
    }
}
