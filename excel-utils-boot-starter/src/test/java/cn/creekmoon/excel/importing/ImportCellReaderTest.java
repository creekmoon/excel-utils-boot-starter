package cn.creekmoon.excel.importing;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.reader.cell.CellReader;
import cn.creekmoon.excel.support.fixture.FixtureCatalog;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** 单元格定位读取（CellReader）：固定坐标、合并单元格、必填校验。 */
@Slf4j
class ImportCellReaderTest {

    /** 冒烟：核心 Happy Path。 */
    @Nested
    class Smoke {

        /** 按 B1/D1/行列坐标读取固定单元格，值与 fixture 一致。 */
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

    /** 分支：合并区、必填失败标记等边界。 */
    @Nested
    class Branch {

        /** 合并单元格读锚点值；引用空白格（A4/F1）应为空串。 */
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
            assertTrue(result.get("A4") == null || result.get("A4").trim().isEmpty());
            assertTrue(result.get("F1") == null || result.get("F1").trim().isEmpty());
            assertEquals(expected.get("A5"), result.get("A5"));
            assertEquals(expected.get("E3"), result.get("E3"));
            log.info("[DONE] merged cell read passed");
        }

        /**
         * mustExist 遇空白单元格：异常被 CellReader 吞掉，仅置 EXISTS_READ_FAIL（不向外抛）。
         */
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
