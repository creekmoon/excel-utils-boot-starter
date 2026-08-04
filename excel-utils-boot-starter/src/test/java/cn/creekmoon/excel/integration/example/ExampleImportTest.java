package cn.creekmoon.excel.integration.example;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.report.ImportGlobalErrorEnum;
import cn.creekmoon.excel.core.R.report.ImportOptions;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.support.assertion.ImportReportAssertions;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.fixture.XlsxTestBuilder;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 复刻 excel-utils-example ExampleController 的导入案例（fixture=导入测试.xlsx）：
 * basic（多 sheet）、with-report（API 报告）、with-range、with-post-processor、with-validation。
 */
class ExampleImportTest {

    private static final String FIXTURE = "导入测试.xlsx";
    private static final int DATA_ROWS = 4;

    @Nested
    class Basic {

        /** 复刻 /import/basic：Sheet1 全 6 列 + Sheet2 部分列（生日/邮箱）。 */
        @Test
        void importBasic_twoSheets_parsesAllAndPartialColumns() throws Exception {
            ExcelImport excelImport = ImportTestHelper.createImport(FIXTURE);

            AtomicInteger sheet1Count = new AtomicInteger();
            List<Student> sheet1 = excelImport.switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvert("全名", Student::setFullName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("邮箱", Student::setEmail)
                    .addConvert("生日", DateConverter::parse, Student::setBirthday)
                    .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                    .read(s -> sheet1Count.incrementAndGet()).getAll();
            assertEquals(DATA_ROWS, sheet1Count.get());
            assertEquals(DATA_ROWS, sheet1.size());

            // 抽样断言：首行/末行字段与解析语义
            Student first = sheet1.get(0);
            assertEquals("zhangsan@qq.com", first.getEmail());
            assertEquals(Integer.valueOf(74), first.getAge());
            assertNotNull(first.getUserName(), "用户名为中文，不应为 null");
            assertNotNull(first.getBirthday());
            assertNotNull(first.getExpTime());
            assertEquals("2023-05-28 17:36:28", new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(first.getBirthday()));

            // Sheet2 只读生日+邮箱两列：邮箱可对上，其他字段为空
            List<Student> sheet2 = excelImport.switchSheet(1, Student::new)
                    .addConvert("生日", DateConverter::parse, Student::setBirthday)
                    .addConvert("邮箱", Student::setEmail)
                    .getAll();
            assertEquals(DATA_ROWS, sheet2.size());
            assertEquals("zhangsan@qq.com", sheet2.get(0).getEmail());
            assertNotNull(sheet2.get(0).getBirthday());
            assertTrue(sheet2.stream().allMatch(s -> s.getAge() == null), "未注册列不应被读取");
        }
    }

    @Nested
    class WithReport {

        /** 复刻 /import/with-report：forApi 预设返回结构化报告。 */
        @Test
        void importWithReport_apiPreset_returnsBalancedReport() throws Exception {
            ImportReport<Student> report = ImportTestHelper.createImport(FIXTURE)
                    .switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("邮箱", Student::setEmail)
                    .addConvert("生日", DateConverter::parse, Student::setBirthday)
                    .readWithReport(ImportOptions.forApi());
            ImportReportAssertions.logReportSummary(report);

            assertEquals(DATA_ROWS, report.getTotalRows());
            assertEquals(DATA_ROWS, report.getSuccessRows());
            assertEquals(0, report.getErrorRows());
            ImportReportAssertions.assertNoErrors(report);
        }
    }

    @Nested
    class WithRange {

        /** 复刻 /import/with-range：range(0, 19) 在小文件上读取全部数据行。 */
        @Test
        void importWithRange_readsAllRowsWithinSmallFile() throws Exception {
            AtomicInteger count = new AtomicInteger();
            ImportTestHelper.createImport(FIXTURE)
                    .switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvert("全名", Student::setFullName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("邮箱", Student::setEmail)
                    .range(0, 19)
                    .read(s -> count.incrementAndGet());
            assertEquals(DATA_ROWS, count.get(), "range 上限超出文件行数时应读完全部数据");
        }
    }

    @Nested
    class WithPostProcessor {

        /** 复刻 /import/with-post-processor：年龄>60 时邮箱必须以 @qq.com 结尾（成功路径）。 */
        @Test
        void postProcessor_validData_allRowsPass() throws Exception {
            ImportReport<Student> report = ImportTestHelper.createImport(FIXTURE)
                    .switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("邮箱", Student::setEmail)
                    .addConvertPostProcessor(student -> {
                        if (student.getAge() != null && student.getAge() > 60) {
                            if (student.getEmail() == null || !student.getEmail().endsWith("@qq.com")) {
                                throw new RuntimeException("年龄大于 60 的用户邮箱必须以 @qq.com 结尾");
                            }
                        }
                    })
                    .readWithReport(ImportOptions.defaults());

            assertEquals(DATA_ROWS, report.getTotalRows());
            assertEquals(DATA_ROWS, report.getSuccessRows());
            ImportReportAssertions.assertNoErrors(report);
        }

        /** 后置处理器抛非自定义异常：该行记错并中断后续读取（IMPORT_INTERRUPTED）。 */
        @Test
        void postProcessor_rejectedRow_interruptsImport() throws Exception {
            Object[][] table = {
                    {"用户名", "年龄", "邮箱"},
                    {"user1", 70, "u1@163.com"},      // 年龄>60 且非 @qq.com → 校验失败
                    {"user2", 30, "u2@qq.com"},
            };
            ImportReport<Student> report = ExcelImport.create(XlsxTestBuilder.multipartFile("pp.xlsx", Arrays.asList(table)))
                    .switchSheet(0, Student::new)
                    .addConvert("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvert("邮箱", Student::setEmail)
                    .addConvertPostProcessor(student -> {
                        if (student.getAge() != null && student.getAge() > 60) {
                            if (student.getEmail() == null || !student.getEmail().endsWith("@qq.com")) {
                                throw new RuntimeException("年龄大于 60 的用户邮箱必须以 @qq.com 结尾");
                            }
                        }
                    })
                    .readWithReport(ImportOptions.defaults());

            assertEquals(1, report.getTotalRows(), "失败行后续不再读取");
            assertEquals(0, report.getSuccessRows());
            assertEquals(1, report.getErrorRows());
            assertTrue(report.getGlobalErrors().contains(ImportGlobalErrorEnum.IMPORT_INTERRUPTED));
        }
    }

    @Nested
    class WithValidation {

        /** 复刻 /import/with-validation：必填+可空+跳过空白组合，fixture 全量通过。 */
        @Test
        void validationRules_allRowsPassOnValidFixture() throws Exception {
            ImportReport<Student> report = ImportTestHelper.createImport(FIXTURE)
                    .switchSheet(0, Student::new)
                    .addConvertAndMustExist("用户名", Student::setUserName)
                    .addConvertAndMustExist("全名", Student::setFullName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .addConvertAndSkipEmpty("邮箱", Student::setEmail)
                    .readWithReport(ImportOptions.defaults());

            assertEquals(DATA_ROWS, report.getTotalRows());
            assertEquals(DATA_ROWS, report.getSuccessRows());
            ImportReportAssertions.assertNoErrors(report);
        }

        /** 必填缺失：该行报「字段[用户名]为必填项」错误。 */
        @Test
        void validationRules_missingRequiredColumn_marksRowError() throws Exception {
            Object[][] table = {
                    {"用户名", "年龄", "邮箱"},
                    {"", 20, "u1@qq.com"},            // 用户名必填缺失
                    {"user2", 30, "u2@qq.com"},
            };
            ImportReport<Student> report = ExcelImport.create(XlsxTestBuilder.multipartFile("req.xlsx", Arrays.asList(table)))
                    .switchSheet(0, Student::new)
                    .addConvertAndMustExist("用户名", Student::setUserName)
                    .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                    .readWithReport(ImportOptions.defaults());

            assertEquals(1, report.getErrorRows());
            assertEquals(1, report.getSuccessRows());
            assertTrue(report.getRowErrors().get(0).getMessage().contains("字段[用户名]为必填项"),
                    report.getRowErrors().get(0).getMessage());
            ImportReportAssertions.assertBalancedStatistics(report);
        }
    }
}
