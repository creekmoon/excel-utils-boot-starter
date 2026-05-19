package cn.creekmoon.excel.importing;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.support.fixture.FixtureCatalog;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.core.R.reader.cell.CellReader;
import cn.hutool.core.util.StrUtil;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

@Slf4j
class ImportCellReaderTest {

    @Nested
    class Smoke {

        @Test
        void importCellReader_fixedCells_valuesMatchTemplate() throws Exception {
            log.info("[STEP] load cell-read fixture={}", FixtureCatalog.CELL_READ_IMPORT);
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.CELL_READ_IMPORT);
            Map<String, String> expected = FixtureCatalog.cellSheet0Expectations();

            AtomicReference<Student> result = new AtomicReference<>();
            excelImport.switchSheetAndUseCellReader(0, Student::new)
                    .addConvert("B1", Student::setUserName)
                    .addConvert("D1", Student::setFullName)
                    .addConvert(0, 5, IntegerConverter::parse, Student::setAge)
                    .read(result::set);

            Student student = result.get();
            log.info("[RESULT] cellRead sheet0: userName={}, fullName={}, age={}",
                    student.getUserName(), student.getFullName(), student.getAge());
            assertEquals(expected.get("B1"), student.getUserName());
            assertEquals(expected.get("D1"), student.getFullName());
            assertEquals(Integer.parseInt(expected.get("age")), student.getAge());
            log.info("[DONE] fixed cell read passed");
        }
    }

    @Nested
    class Branch {

        @Test
        void importCellReader_mergedCells_readsAnchorAndBlankReferenceCells() throws Exception {
            log.info("[STEP] read merged cells on sheet1");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.CELL_READ_IMPORT);
            Map<String, String> expected = FixtureCatalog.cellSheet1Expectations();

            AtomicReference<Map<String, String>> holder = new AtomicReference<>();
            excelImport.switchSheetAndUseCellReader(1, () -> new HashMap<String, String>())
                    .addConvert("A1", (x, y) -> x.put("A1", y))
                    .addConvert("E1", (x, y) -> x.put("E1", y))
                    .addConvert("A4", (x, y) -> x.put("A4", y))
                    .addConvert("F1", (x, y) -> x.put("F1", y))
                    .addConvert("A5", (x, y) -> x.put("A5", y))
                    .addConvert("E3", (x, y) -> x.put("E3", y))
                    .read(holder::set);

            Map<String, String> result = holder.get();
            log.info("[RESULT] merged cell map={}", result);
            assertEquals(expected.get("A1"), result.get("A1"));
            assertEquals(expected.get("E1"), result.get("E1"));
            assertTrue(StrUtil.isBlank(result.get("A4")));
            assertTrue(StrUtil.isBlank(result.get("F1")));
            assertEquals(expected.get("A5"), result.get("A5"));
            assertEquals(expected.get("E3"), result.get("E3"));
            log.info("[DONE] merged cell read passed");
        }

        @Test
        void importCellReader_mustExistBlankCell_marksReadFailure() throws Exception {
            log.info("[STEP] configure mustExist on blank cell A4 (sheet1 merged area)");
            ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.CELL_READ_IMPORT);
            AtomicReference<Student> holder = new AtomicReference<>();

            CellReader<Student> reader = excelImport.switchSheetAndUseCellReader(1, Student::new)
                    .addConvert("A1", Student::setUserName)
                    .addConvertAndMustExist("A4", Student::setFullName);
            reader.read(holder::set);

            log.info("[RESULT] read completed, existsReadFail={}", reader.EXISTS_READ_FAIL.get());
            assertTrue(reader.EXISTS_READ_FAIL.get());
            log.info("[DONE] mustExist blank cell marks read failure without propagating exception");
        }
    }
}
