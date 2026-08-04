package cn.creekmoon.excel.importing;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.reader.cell.CellReader;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.XlsxTestBuilder;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** CellReader 行为边界：选填跳过、必填成功、成功路径无失败标记。 */
class ImportCellReaderBehaviorTest {

    /** 第 0 行 B1 单元格放文本 "B1"，第 1 行 C2 单元格放文本 "C2"，供坐标引用断言。 */
    private static final Object[][] TABLE = {
            {"A", "B1", "C", "D", "E"},
            {"A2", "B2", "C2", "D2", "E2"},
            {"A3", "B3", "C3", "D3", "E3"},
            {"A4", "B4", "C4", "D4", "E4"},
    };

    /** skipEmpty 空白单元格：字段不被设置，行仍成功。 */
    @Test
    void skipEmpty_blankCell_fieldNotSet_noFailure() throws Exception {
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("cell-skip.xlsx", Arrays.asList(TABLE)));
        AtomicReference<Student> holder = new AtomicReference<>();

        // E2 在数据表中是 "E2" 非空；B1 正常读取；引用一个超出表格的空白格（H10）验证跳过
        CellReader<Student> reader = excelImport.switchSheetAndUseCellReader(0, Student::new)
                .addConvert("B1", Student::setUserName)
                .addConvertAndSkipEmpty("H10", Student::setFullName);
        reader.read(holder::set);

        Student student = holder.get();
        assertEquals("B1", student.getUserName());
        assertNull(student.getFullName(), "skipEmpty 空白单元格不应设置字段");
        assertFalse(reader.EXISTS_READ_FAIL.get(), "成功路径不应有读取失败标记");
    }

    /** 空白单元格不加 skipEmpty 也无 mustExist：按空字符串转换，不报错。 */
    @Test
    void blankCell_withoutRules_convertWithEmptyString() throws Exception {
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("cell-plain.xlsx", Arrays.asList(TABLE)));
        AtomicReference<Student> holder = new AtomicReference<>();

        CellReader<Student> reader = excelImport.switchSheetAndUseCellReader(0, Student::new)
                .addConvert("B1", Student::setUserName)
                .addConvert("H10", Student::setFullName);
        reader.read(holder::set);

        Student student = holder.get();
        assertEquals("B1", student.getUserName());
        assertEquals("", student.getFullName(), "无规则空白单元格按空字符串转换");
        assertFalse(reader.EXISTS_READ_FAIL.get());
    }

    /** mustExist 有值单元格：正常读取且无失败标记。 */
    @Test
    void mustExist_nonBlankCell_readsNormally() throws Exception {
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("cell-must.xlsx", Arrays.asList(TABLE)));
        AtomicReference<Student> holder = new AtomicReference<>();

        CellReader<Student> reader = excelImport.switchSheetAndUseCellReader(0, Student::new)
                .addConvertAndMustExist("B1", Student::setUserName)
                .addConvertAndMustExist("C2", Student::setFullName);
        reader.read(holder::set);

        Student student = holder.get();
        assertEquals("B1", student.getUserName());
        assertEquals("C2", student.getFullName());
        assertFalse(reader.EXISTS_READ_FAIL.get());
    }

    /** 数字单元格按 StringConverter 转成字符串（如 "2000.0"），再走用户转换器。 */
    @Test
    void numericCell_convertedToPlainString() throws Exception {
        Object[][] numericTable = {
                {2000.0, "文本"},
        };
        ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("cell-num.xlsx", Arrays.asList(numericTable)));
        AtomicReference<Student> holder = new AtomicReference<>();

        CellReader<Student> reader = excelImport.switchSheetAndUseCellReader(0, Student::new)
                .addConvert("A1", (v) -> "prefix:" + v, Student::setUserName);
        reader.read(holder::set);

        assertEquals("prefix:2000", holder.get().getUserName());
        assertFalse(reader.EXISTS_READ_FAIL.get());
    }
}
