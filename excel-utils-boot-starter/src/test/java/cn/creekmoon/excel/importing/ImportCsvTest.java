package cn.creekmoon.excel.importing;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.assertion.PoiWorkbookAssertions;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

@Slf4j
class ImportCsvTest {

    private static final int EXPECTED_CSV_DATA_ROWS = 3;

    @Test
    void importCsv_convertToXlsxThenRead_rowCountAndFieldsMatchSource() throws Exception {
        MockMultipartFile csvFile = ImportTestHelper.loadFixture("import-demo.csv");
        String csvText = new String(csvFile.getBytes(), StandardCharsets.UTF_8);
        long csvDataRows = csvText.lines().skip(1).filter(line -> !line.isBlank()).count();
        log.info("[STEP] csv fixture loaded, header+data lines={}, expectedDataRows={}", csvText.lines().count(), csvDataRows);

        MockMultipartFile multipartFile = new MockMultipartFile(
                "import-demo.csv", "import-demo.csv", "text/csv", csvFile.getBytes());

        ExcelImport excelImport = ExcelImport.create(multipartFile);
        TitleReader<Student> reader = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail);

        List<Student> students = reader.getAll();
        ImportReport<Student> report = reader.getReport();
        ImportReportAssertions.logReportSummary(report);

        assertEquals(EXPECTED_CSV_DATA_ROWS, csvDataRows);
        assertEquals(EXPECTED_CSV_DATA_ROWS, students.size());
        ImportReportAssertions.assertNoErrors(report);

        for (int i = 0; i < EXPECTED_CSV_DATA_ROWS; i++) {
            Student student = students.get(i);
            log.info("[ASSERT] csv row {}: userName={}, email={}, age={}", i, student.getUserName(), student.getEmail(), student.getAge());
            assertEquals("user" + i, student.getUserName());
            assertEquals("full" + i, student.getFullName());
            assertEquals(20 + i, student.getAge());
            assertEquals("user" + i + "@qq.com", student.getEmail());
        }

        try (Workbook workbook = PoiWorkbookAssertions.open(
                new java.io.ByteArrayInputStream(excelImport.sourceFile.getInputStream().readAllBytes()))) {
            PoiWorkbookAssertions.assertRowCount(workbook.getSheetAt(0), EXPECTED_CSV_DATA_ROWS + 1);
            log.info("[ASSERT] converted xlsx rowCount={}", workbook.getSheetAt(0).getLastRowNum() + 1);
        }
        log.info("[DONE] csv import passed");
    }
}
