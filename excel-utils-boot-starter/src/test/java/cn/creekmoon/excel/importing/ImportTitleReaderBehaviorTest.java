package cn.creekmoon.excel.importing;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.report.ImportGlobalErrorEnum;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.XlsxTestBuilder;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * TitleReader 行为边界：必填/选填、转换失败、空白行过滤、模板检查开关、列序无关性、空 sheet。
 * 全部使用动态生成的 xlsx 字节构造受控输入。
 */
class ImportTitleReaderBehaviorTest {

    private static final Object[][] STD_TABLE = {
            {"用户名", "年龄", "邮箱"},
            {"user1", 20, "u1@qq.com"},
            {"user2", "abc", "u2@qq.com"},          // 年龄非法
            {"user3", 30, ""},                       // 邮箱为空
            {"", 40, "u4@qq.com"},                   // 用户名为空（空白字符串行）
            {"user5", 50, "u5@qq.com"},
    };

    private TitleReader<Student> bindStd(TitleReader<Student> reader) {
        return reader
                .addConvert("用户名", Student::setUserName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail);
    }

    @Nested
    class ValidationRules {

        /** 年龄列空值 + mustExist：该行报「字段[年龄]为必填项」，其余行正常。 */
        @Test
        void mustExist_blankValue_marksRowErrorWithFieldMessage() throws Exception {
            Object[][] table = {
                    {"用户名", "年龄", "邮箱"},
                    {"user1", 20, "u1@qq.com"},
                    {"user2", "", "u2@qq.com"},      // 年龄必填缺失
                    {"user3", 30, "u3@qq.com"},
            };
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("must.xlsx", rows(table)));
            ImportReport report = excelImport.switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvertAndMustExist("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("邮箱", Student::setEmail)
                    .readWithReport();

            assertEquals(1, report.getErrorRows());
            assertTrue(report.getRowErrors().get(0).getMessage().contains("字段[年龄]"));
            assertEquals(2, report.getSuccessRows());
            ImportReportAssertions.assertBalancedStatistics(report);
        }

        /** 空值 + skipEmpty：字段被跳过不报错，其余字段正常解析。 */
        @Test
        void skipEmpty_blankValue_skipsFieldWithoutError() throws Exception {
            Object[][] table = {
                    {"用户名", "年龄", "邮箱"},
                    {"user1", 20, "u1@qq.com"},
                    {"user2", "", "u2@qq.com"},      // 年龄选填缺失
                    {"user3", 30, ""},               // 邮箱选填缺失
                    {"user4", 40, "u4@qq.com"},
            };
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("skip.xlsx", rows(table)));
            List<Student> students = excelImport.switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvertAndSkipEmpty("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvertAndSkipEmpty("邮箱", Student::setEmail)
                    .getAll();

            assertEquals(4, students.size());
            Student noAge = students.stream().filter(s -> "user2".equals(s.getUserName())).findFirst().orElseThrow();
            assertNull(noAge.getAge());
            Student noEmail = students.stream().filter(s -> "user3".equals(s.getUserName())).findFirst().orElseThrow();
            assertNull(noEmail.getEmail());
            assertNotNull(noEmail.getAge());
        }

        @Test
        void convertFailure_rowErrorContainsColumnName() throws Exception {
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("conv.xlsx", rows(STD_TABLE)));
            ImportReport report = bindStd(excelImport.switchSheet(0, Student::new)).readWithReport();

            assertEquals(1, report.getErrorRows());
            assertTrue(report.getRowErrors().get(0).getMessage().contains("字段[年龄]解析失败"), report.getRowErrors().get(0).getMessage());
            assertEquals(4, report.getSuccessRows());
            ImportReportAssertions.assertBalancedStatistics(report);
        }
    }

    @Nested
    class BlankRowFilter {

        /** 全空字符串行（3 列均空）：默认被空白行过滤，2 条有效数据计入统计。 */
        @Test
        void fullyBlankRow_notCounted_byDefault() throws Exception {
            Object[][] withEmpty = {
                    {"用户名", "年龄", "邮箱"},
                    {"user1", 20, "u1@qq.com"},
                    {"", "", ""},
                    {"user2", 30, "u2@qq.com"},
            };
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("blank.xlsx", rows(withEmpty)));
            ImportReport report = bindStd(excelImport.switchSheet(0, Student::new)).readWithReport();

            assertEquals(2, report.getTotalRows());
            assertEquals(2, report.getSuccessRows());
            ImportReportAssertions.assertNoErrors(report);
        }

        /**
         * 现状契约：全空字符串行在 EasyExcel SAX 层即被跳过（不产生 invoke），
         * disableBlankRowFilter 无法恢复——过滤开关在 EasyExcel 引擎下只对
         * 「SAX 层存在但全部源列解析为空」的行生效，而这类行实际不存在。
         */
        @Test
        void fullyBlankRow_notCounted_evenWhenFilterDisabled() throws Exception {
            Object[][] withEmpty = {
                    {"用户名", "年龄", "邮箱"},
                    {"user1", 20, "u1@qq.com"},
                    {"", "", ""},
                    {"user2", 30, "u2@qq.com"},
            };
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("blank-off.xlsx", rows(withEmpty)));
            cn.creekmoon.excel.core.R.reader.title.EasyExcelTitleReader<Student> reader =
                    (cn.creekmoon.excel.core.R.reader.title.EasyExcelTitleReader<Student>)
                            excelImport.switchSheet(0, Student::new);
            reader.addConvert("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("邮箱", Student::setEmail);
            ImportReport report = reader.disableBlankRowFilter().readWithReport();

            assertEquals(2, report.getTotalRows(), "全空字符串行在 SAX 层不产生 invoke");
        }

        /**
         * 行只有未注册列（备注）有值时：colIndex2Title 含全部源列，行不视为空白；
         * 注册的空值列按空串转换，IntegerConverter.parse("") 失败 → 该行记错。
         */
        @Test
        void rowWithOnlyUnregisteredColumnValue_partialBlankConvertsWithError() throws Exception {
            Object[][] withRemark = {
                    {"用户名", "年龄", "邮箱", "备注"},
                    {"user1", 20, "u1@qq.com", "x"},
                    {"", "", "", "只有备注列有值"},
                    {"user2", 30, "u2@qq.com", "y"},
            };
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("remark.xlsx", rows(withRemark)));
            ImportReport report = bindStd(excelImport.switchSheet(0, Student::new)).readWithReport();

            assertEquals(3, report.getTotalRows(), "有未注册列值的行不被空白过滤");
            assertEquals(2, report.getSuccessRows());
            assertEquals(1, report.getErrorRows(), "空年龄按空串转换失败");
            ImportReportAssertions.assertBalancedStatistics(report);
        }
    }

    @Nested
    class TemplateConsistency {

        /** 缺失列：模板不匹配属于文件级失败，立即停止且不扫描数据行。 */
        @Test
        void missingColumn_marksTemplateMismatch_stopsImmediately() throws Exception {
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("tmpl.xlsx", rows(STD_TABLE)));
            ImportReport report = excelImport.switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvert("不存在的列", Student::setEmail)
                    .readWithReport();

            assertTrue(report.getGlobalErrors().contains(ImportGlobalErrorEnum.TEMPLATE_MISMATCH));
            assertEquals(0, report.getTotalRows());
            assertEquals(0, report.getSuccessRows());
            assertEquals(0, report.getErrorRows());
        }

        /** disableTemplateConsistencyCheck：缺失列不产生全局错误，可正常读取。 */
        @Test
        void disabledTemplateCheck_ignoresMissingColumn() throws Exception {
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("tmpl-off.xlsx", rows(STD_TABLE)));
            // disableTemplateConsistencyCheck 是 EasyExcelTitleReader 的方法，经门面返回的 TitleReader 需强转
            cn.creekmoon.excel.core.R.reader.title.EasyExcelTitleReader<Student> reader =
                    (cn.creekmoon.excel.core.R.reader.title.EasyExcelTitleReader<Student>)
                            excelImport.switchSheet(0, Student::new);
            reader.addConvert("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("邮箱", Student::setEmail)
                    .addConvert("不存在的列", Student::setEmail);
            ImportReport report = reader.disableTemplateConsistencyCheck().readWithReport();

            assertFalse(report.getGlobalErrors().contains(ImportGlobalErrorEnum.TEMPLATE_MISMATCH));
            assertEquals(5, report.getTotalRows());
            assertEquals(4, report.getSuccessRows(), "abc 行转换失败，成功 4 行");
            assertEquals(1, report.getErrorRows());
        }
    }

    @Nested
    class ColumnOrderIndependence {

        /** 表头列乱序 + 多余列：按标题映射取值，与声明顺序无关。 */
        @Test
        void shuffledAndExtraColumns_valuesMapByTitle() throws Exception {
            Object[][] shuffled = {
                    {"邮箱", "年龄", "用户名", "多余列"},
                    {"u1@qq.com", 20, "user1", "X"},
                    {"u2@qq.com", 21, "user2", "Y"},
            };
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("shuffled.xlsx", rows(shuffled)));
            List<Student> students = bindStd(excelImport.switchSheet(0, Student::new)).getAll();

            assertEquals(2, students.size());
            Student s0 = students.get(0);
            assertEquals("user1", s0.getUserName());
            assertEquals(20, s0.getAge());
            assertEquals("u1@qq.com", s0.getEmail());
        }

        /** 只注册部分列：未注册列忽略，不报错。 */
        @Test
        void partialColumns_unregisteredIgnored() throws Exception {
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("partial.xlsx", rows(STD_TABLE)));
            List<Student> students = excelImport.switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .getAll();

            assertEquals(5, students.size());
            assertTrue(students.stream().allMatch(s -> s.getUserName() != null));
            assertTrue(students.stream().allMatch(s -> s.getAge() == null));
        }
    }

    @Nested
    class EmptyAndEdgeSheets {

        @Test
        void headerOnlySheet_totalRowsZero() throws Exception {
            // 表头列需配齐绑定列，避免模板不匹配全局错误干扰"空数据"语义
            Object[][] headerOnly = {{"用户名", "年龄", "邮箱"}};
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("header-only.xlsx", rows(headerOnly)));
            ImportReport report = bindStd(excelImport.switchSheet(0, Student::new)).readWithReport();

            assertEquals(0, report.getTotalRows());
            assertEquals(0, report.getSuccessRows());
            assertEquals(0, report.getErrorRows());
            ImportReportAssertions.assertNoErrors(report);
        }

        @Test
        void noConvertsRegistered_skipsAllDataRows() throws Exception {
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("no-conv.xlsx", rows(STD_TABLE)));
            ImportReport report = excelImport.switchSheet(0, Student::new).readWithReport();

            assertEquals(0, report.getTotalRows());
        }

        @Test
        void reportMetadata_populatedFromConfig() throws Exception {
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("meta.xlsx", rows(STD_TABLE)));
            ImportReport report = bindStd(excelImport.switchSheet(0, Student::new)).readWithReport();

            assertNotNull(report.getTaskId());
            assertEquals("Sheet0", report.getSheetName());
            assertNotNull(report.getSheetRid());
            assertEquals(0, report.getTitleRowIndex());
            assertEquals(1, report.getFirstRowIndex());
            assertEquals(Integer.MAX_VALUE, report.getLatestRowIndex());
            // 实际处理首尾行（含 1 行转换失败，共 5 行数据）
            assertEquals(Integer.valueOf(1), report.getProcessedFirstRowIndex());
            assertEquals(Integer.valueOf(5), report.getProcessedLastRowIndex());
        }
    }

    @Nested
    class RepeatedRead {

        /** 每次 read 都重置报告，重复读取不会让统计累加。 */
        @Test
        void secondRead_resetsReport() throws Exception {
            ExcelImport excelImport = ExcelImport.create(XlsxTestBuilder.multipartFile("repeat.xlsx", rows(STD_TABLE)));
            TitleReader<Student> reader = bindStd(excelImport.switchSheet(0, Student::new));

            reader.read(x -> {});
            ImportReport first = reader.getReport();
            assertEquals(5, first.getTotalRows());

            reader.read(x -> {});
            ImportReport second = reader.getReport();
            assertEquals(5, second.getTotalRows(), "第二次 read 应重置报告，统计不累加");
            assertEquals(4, second.getSuccessRows());
            assertEquals(1, second.getErrorRows());
        }
    }

    /** 表头行 + 数据行构建器（Object[][] 转 List<Object[]>，不能直接用 List.of）。 */
    private List<Object[]> rows(Object[][] table) {
        return Arrays.asList(table);
    }
}
