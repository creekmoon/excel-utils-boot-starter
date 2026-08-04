package cn.creekmoon.excel.importing;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.XlsxTestBuilder;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.List;
import java.util.function.Supplier;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** ExcelImport 门面异常路径与 CSV 分流契约。 */
class ImportFacadeTest {

    private static final Object[][] TABLE = {
            {"用户名", "年龄"},
            {"user1", 20},
            {"user2", 30},
    };

    @Test
    void switchSheet_negativeIndex_throws() throws Exception {
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("neg.xlsx", Arrays.asList(TABLE)));
        assertThrows(RuntimeException.class, () -> excelImport.switchSheet(-1, Student::new));
    }

    @Test
    void switchSheet_indexOutOfBounds_throws() throws Exception {
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("oob.xlsx", Arrays.asList(TABLE)));
        assertThrows(RuntimeException.class, () -> excelImport.switchSheet(99, Student::new));
    }

    @Test
    void switchSheetAndUseCellReader_indexOutOfBounds_throws() throws Exception {
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("oob2.xlsx", Arrays.asList(TABLE)));
        assertThrows(RuntimeException.class, () -> excelImport.switchSheetAndUseCellReader(1, Student::new));
    }

    /** supplier 两次返回同一实例：触发「不支持重写 equal()」断言（现状锁定）。 */
    @Test
    void switchSheet_cachedSupplierInstance_throws() throws Exception {
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("cache.xlsx", Arrays.asList(TABLE)));
        Supplier<Student> cached = new Supplier<>() {
            private final Student s = new Student();

            @Override
            public Student get() {
                return s;
            }
        };
        RuntimeException e = assertThrows(RuntimeException.class, () -> excelImport.switchSheet(0, cached));
        assertTrue(e.getMessage().contains("equal"), e.getMessage());
    }

    /** 正常 supplier（每次 new）不触发断言。 */
    @Test
    void switchSheet_freshSupplierInstance_ok() throws Exception {
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("fresh.xlsx", Arrays.asList(TABLE)));
        TitleReader<Student> reader = excelImport.switchSheet(0, Student::new);
        assertNotNull(reader);
    }

    /** 非 Excel 文件（文本内容）：create 时解析 workbook.xml 失败抛出。 */
    @Test
    void create_nonExcelFile_throws() {
        MockMultipartFile badFile = new MockMultipartFile(
                "bad.xlsx", "bad.xlsx", null, "这不是Excel文件".getBytes(StandardCharsets.UTF_8));
        assertThrows(RuntimeException.class, () -> ExcelImport.create(badFile));
    }

    /** 空文件：同样无法解析 Sheet 映射。 */
    @Test
    void create_emptyFile_throws() {
        MockMultipartFile empty = new MockMultipartFile("empty.xlsx", "empty.xlsx", null, new byte[0]);
        assertThrows(RuntimeException.class, () -> ExcelImport.create(empty));
    }

    /** CSV 扩展名大小写不敏感：大写 .CSV 同样触发转 xlsx。 */
    @Test
    void csvSupport_uppercaseExtension_convertsToXlsx() throws Exception {
        String csv = "用户名,年龄\nuser1,20\nuser2,30\n";
        MockMultipartFile csvFile = new MockMultipartFile(
                "DATA.CSV", "DATA.CSV", "text/csv", csv.getBytes(StandardCharsets.UTF_8));

        ExcelImport excelImport = ExcelImport.create(csvFile);
        // MockMultipartFile(name, inputStream) 构造器不设 originalFilename，用 getName 断言转换已发生
        assertEquals("csv2xlsx.xlsx", excelImport.sourceFile.getName(),
                "大写 .CSV 也应触发内部转 xlsx");
    }

    /** 非 CSV（xlsx）：不触发转换，sourceFile 保持原样。 */
    @Test
    void csvSupport_xlsxFile_notConverted() throws Exception {
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("plain.xlsx", Arrays.asList(TABLE)));
        assertEquals("plain.xlsx", excelImport.sourceFile.getOriginalFilename());
    }
}
