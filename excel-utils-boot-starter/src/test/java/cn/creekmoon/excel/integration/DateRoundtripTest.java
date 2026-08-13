package cn.creekmoon.excel.integration;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.StudentFixtures;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 日期列全链路契约：导出 Date / LocalDateTime 列 → 再导入解析，值保持到秒级。
 * 这是导入导出组件的核心卖点，旧 roundtrip 只覆盖了数值列。
 */
class DateRoundtripTest {

    private static final int ROW_COUNT = 10;

    @Test
    void roundtrip_dateAndLocalDateTimeColumns_valuesPreservedToSecond() throws Exception {
        ArrayList<Student> data = StudentFixtures.createList(ROW_COUNT);

        byte[] bytes = exportWithDateColumns(data);

        TitleReader<Student> reader = ExcelImport.create(
                        new MockMultipartFile("dates.xlsx", new ByteArrayInputStream(bytes)))
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime);
        List<Student> imported = reader.getAll();
        ImportReport report = reader.getReport();
        ImportReportAssertions.logReportSummary(report);
        ImportReportAssertions.assertNoErrors(report);

        assertEquals(ROW_COUNT, imported.size());
        for (int i = 0; i < ROW_COUNT; i++) {
            Student src = data.get(i);
            Student dst = imported.get(i);
            assertEquals(truncateToSeconds(src.getBirthday()), truncateToSeconds(dst.getBirthday()),
                    "birthday row " + i);
            assertEquals(src.getExpTime().withNano(0), dst.getExpTime(), "expTime row " + i);
        }
    }

    @Test
    void roundtrip_nullDateColumns_skippedBySkipEmpty() throws Exception {
        ArrayList<Student> data = StudentFixtures.createList(2);
        data.get(1).setBirthday(null);
        data.get(1).setExpTime(null);

        byte[] bytes = exportWithDateColumns(data);

        TitleReader<Student> reader = ExcelImport.create(
                        new MockMultipartFile("null-dates.xlsx", new ByteArrayInputStream(bytes)))
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvertAndSkipEmpty("生日", DateConverter::parse, Student::setBirthday)
                .addConvertAndSkipEmpty("过期时间", LocalDateTimeConverter::parse, Student::setExpTime);
        List<Student> imported = reader.getAll();
        ImportReportAssertions.assertNoErrors(reader.getReport());

        assertEquals(2, imported.size());
        assertNull(imported.get(1).getBirthday(), "null 生日列应被 skipEmpty 跳过");
        assertNull(imported.get(1).getExpTime(), "null 过期时间列应被 skipEmpty 跳过");
        assertTrue(imported.get(0).getBirthday() != null, "非空日期列正常解析");
    }

    private byte[] exportWithDateColumns(ArrayList<Student> data) throws Exception {
        ExcelExport excelExport = ExcelExport.create();
        excelExport.switchNewSheet(Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(data);
        return Files.readAllBytes(excelExport.stopWrite().toPath());
    }

    private static Date truncateToSeconds(Date date) {
        return new Date((date.getTime() / 1000) * 1000);
    }
}
