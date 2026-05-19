package cn.creekmoon.excel.integration;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.assertion.PoiWorkbookAssertions;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.StudentFixtures;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.jupiter.api.Assertions.assertEquals;

@Slf4j
class ImportExportRoundtripTest {

    @Test
    void exportThenImport_roundtrip_preservesRowCountAndSampleFields() throws Exception {
        int targetCount = 600;
        ArrayList<Student> data = StudentFixtures.createList(targetCount);
        log.info("[STEP] export {} deterministic students", targetCount);

        ExcelExport excelExport = ExcelExport.create();
        excelExport.switchNewSheet(Student.class)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .write(data);
        byte[] bytes = Files.readAllBytes(excelExport.stopWrite().toPath());
        log.info("[RESULT] exported bytes={}", bytes.length);

        try (Workbook exportWorkbook = PoiWorkbookAssertions.open(new ByteArrayInputStream(bytes))) {
            Sheet sheet = exportWorkbook.getSheetAt(0);
            PoiWorkbookAssertions.assertRowCount(sheet, targetCount + 1);
            PoiWorkbookAssertions.assertCellString(sheet, 1, 0, String.valueOf(data.get(0).getAge()));
            PoiWorkbookAssertions.assertCellString(sheet, 1, 1, data.get(0).getEmail());
        }

        log.info("[STEP] import exported bytes and verify rows");
        AtomicInteger count = new AtomicInteger();
        TitleReader<Student> reader = ExcelImport.create(new MockMultipartFile("roundtrip.xlsx", new ByteArrayInputStream(bytes)))
                .switchSheet(0, Student::new)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail);
        List<Student> imported = reader.read(item -> count.incrementAndGet()).getAll();
        ImportReport<Student> report = reader.getReport();
        ImportReportAssertions.logReportSummary(report);
        ImportReportAssertions.assertNoErrors(report);

        assertEquals(targetCount, count.get());
        assertEquals(targetCount, imported.size());
        assertEquals(data.get(0).getAge(), imported.get(0).getAge());
        assertEquals(data.get(0).getEmail(), imported.get(0).getEmail());
        assertEquals(data.get(targetCount - 1).getAge(), imported.get(targetCount - 1).getAge());
        assertEquals(data.get(targetCount - 1).getEmail(), imported.get(targetCount - 1).getEmail());
        log.info("[ASSERT] first row age/email={}, {} | last row age/email={}, {}",
                imported.get(0).getAge(), imported.get(0).getEmail(),
                imported.get(targetCount - 1).getAge(), imported.get(targetCount - 1).getEmail());
        log.info("[DONE] export-import roundtrip passed");
    }
}
