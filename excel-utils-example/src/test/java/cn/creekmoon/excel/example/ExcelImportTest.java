package cn.creekmoon.excel.example;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.ExcelFileUtils;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ResourceUtil;
import cn.hutool.core.util.StrUtil;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.mock.web.MockHttpSession;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.ResultActions;
import org.springframework.test.web.servlet.request.MockMvcRequestBuilders;
import org.springframework.test.web.servlet.result.MockMvcResultMatchers;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.*;

@AutoConfigureMockMvc
@SpringBootTest
@Slf4j
class ExcelImportTest {

    @Autowired
    private MockMvc mvc;
    private MockHttpSession session;

    @BeforeEach
    public void setupMockMvc() {
        session = new MockHttpSession();
    }

    /**
     * 测试Response方法
     *
     * @throws Exception
     */
    @Test
    public void controllerResponseTest() throws Exception {
        AtomicInteger count = new AtomicInteger(0);
        int targetCount = 600;
        /*调用导出Controller*/
        ResultActions resultActions = this.mvc.perform(MockMvcRequestBuilders
                .get(URI.create("/exampleTest"))
                .queryParam("size", String.valueOf(targetCount)));

        /*断言导出结果*/
        resultActions.andReturn().getResponse().setCharacterEncoding("UTF-8");
        resultActions.andExpect(MockMvcResultMatchers.status().isOk())
                .andDo(x -> {
                    MockMultipartFile mockFile = new MockMultipartFile("mockFile", x.getResponse().getContentAsByteArray());
                    ExcelImport.create(mockFile)
                            .switchSheet(0, Student::new)
                            .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                            .addConvert("邮箱", Student::setEmail)
                            .read(data -> {
                                log.info("[测试导出] Object={}", data);
                                count.incrementAndGet();
                            });
                });

        // 断言返回值是600行
        Assertions.assertEquals(targetCount, count.get());
    }


    /**
     * 测试临时文件清理
     *
     * @throws Exception
     */
    @Test
    public void tempFileTest() throws Exception {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);


        /*计数器*/
        AtomicInteger count = new AtomicInteger();

        /*第一个sheet导入测试*/
        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        TitleReader<Student> read = excelImport
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .read(student -> {
                    log.info("[测试导入] object={}", student);
                    count.getAndIncrement();
                });
        Assertions.assertEquals(1000, count.get());

        /*第二个sheet导入测试*/
        List<Student> students = excelImport.switchSheet(1, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .getAll();
        Assertions.assertEquals(4, students.size());

        //检查临时文件是否能够正常生成和清理
        File file = excelImport.generateResultFile();
        Assertions.assertTrue(FileUtil.exist(file));
        ExcelFileUtils.cleanTempFileByPathNow(file.getPath());
        Assertions.assertFalse(FileUtil.exist(file));

    }


    /**
     * 测试读取指定范围
     *
     * @throws Exception
     */
    @Test
    public void rangeTest() throws Exception {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);

        /*导入测试*/
        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        TitleReader<Student> studentSheetReader = excelImport
                .switchSheet(0, Student::new)
                .range(0, 3, 4)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime);
        List<Student> students = studentSheetReader.getAll();

        /*检查是否能够正确读取*/
        Assertions.assertEquals(2, students.size());
    }


    /**
     * 单元格默认读取
     *
     * @throws Exception
     */
    @Test
    public void cellReadTest() throws Exception {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-cell-read.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);


        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        // 按index和 按单元格名称读取测试
        AtomicReference<Student> sheet1 = new AtomicReference<>();
        excelImport
                .switchSheetAndUseCellReader(0, Student::new)
                .addConvert("B1", Student::setUserName)
                .addConvert("D1", Student::setFullName)
                .addConvert(0, 5, IntegerConverter::parse, Student::setAge)
                .read(sheet1::set);
        Assertions.assertEquals(sheet1.get().getUserName(), "李二狗");
        Assertions.assertEquals(sheet1.get().getFullName(), "李云龙");
        Assertions.assertEquals(sheet1.get().getAge(), 2000);


        // 读取合并单元格测试
        excelImport
                .switchSheetAndUseCellReader(1, HashMap::new)
                .addConvert("A1", (x, y) -> x.put("A1", y))
                .addConvert("E1", (x, y) -> x.put("E1", y))
                .addConvert("A4", (x, y) -> x.put("A4", y))
                .addConvert("F1", (x, y) -> x.put("F1", y))
                .addConvert("A5", (x, y) -> x.put("A5", y))
                .addConvert("E3", (x, y) -> x.put("E3", y))
                .read(x -> {
                    Assertions.assertEquals(x.get("A1"), "合并单元格(A1)");
                    Assertions.assertEquals(x.get("E1"), "合并单元格(E1)");
                    Assertions.assertTrue(StrUtil.isBlank((String) x.get("A4")));
                    Assertions.assertTrue(StrUtil.isBlank((String) x.get("F1")));
                    Assertions.assertEquals(x.get("A5"), "合并单元格(A5)");
                    Assertions.assertEquals(x.get("E3"), "合并单元格(E3)");
                    System.out.println("测试通过 x = " + x);
                });

    }


    @Test
    void importTest() throws IOException, InterruptedException {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);

        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        TitleReader<Student> read = excelImport.switchSheet(0, Student::new);
        List<Student> dataList  = read
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .getAll();

        // 数据从第二行开始, 而索引下标从0开始, 所以需要都-2
        assertEquals(88, dataList.get(6 - 2).age, "第6行数据年龄为88");
        assertEquals("poxo0", dataList.get(80 - 2).getUserName(), "第80行数据用户名是poxo0");
        assertEquals("npdsytu8i5@qq.com", dataList.get(195 - 2).getEmail(), "第195行数据邮箱是npdsytu8i5@qq.com");
        //第630条全名是j94rk
        assertEquals("j94rk", dataList.get(631 - 2).getFullName(), "第631行数据全名是j94rk");
        //全部数据有1000条
        assertEquals(1000, dataList.size(), "全部数据有1000条");
        
        // 大版本升级：使用 ImportReport 替代 rowIndex2msg
        assertEquals(1, read.getReport().getProcessedFirstRowIndex(), "数据起始行下标预期为1");
        assertEquals(1000, read.getReport().getProcessedLastRowIndex(), "数据结束行下标预期为1000");
        assertEquals(1000, read.getReport().getSuccessRows(), "成功行数预期为1000");

        long countAgeLg60 = dataList
                .stream()
                .map(Student::getAge)
                .filter(x -> x > 60)
                .count();
        assertEquals(countAgeLg60, 411, "年龄大于60应该为411个");

    }


    @Test
    void importTest2() throws IOException, InterruptedException {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);

        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .getAll();

        TitleReader<Student> sheet2 = excelImport.switchSheet(1, Student::new)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("邮箱", Student::setEmail)
                .read(x -> {
                    if ("lisi@qq.com".equals(x.getEmail())) {
                        Thread.sleep(2000);
                        throw new RuntimeException("错误!");
                    }
                });
        // 大版本升级：使用 ImportReport 验证错误
        assertEquals(1, sheet2.getReport().getErrorRows(), "预期发生一个错误");

        File file = excelImport.generateResultFile();
        assertTrue(FileUtil.exist(file), "文件预期应该正常生成");

    }

    /**
     * 测试模板不匹配 + fail-fast
     */
    @Test
    void importReportTemplateMismatchTest() throws Exception {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);

        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        
        // 使用错误的标题配置（故意让模板不匹配）
        cn.creekmoon.excel.core.R.report.ImportOptions options = new cn.creekmoon.excel.core.R.report.ImportOptions();
        options.setFailFastOnTemplateMismatch(true);
        
        cn.creekmoon.excel.core.R.report.ImportReport<Student> report = excelImport
                .switchSheet(0, Student::new)
                .addConvert("不存在的标题", Student::setUserName)
                .readWithReport(options);
        
        // 验证模板错误
        assertEquals("TEMPLATE_MISMATCH", report.getGlobalErrorCode(), "应该检测到模板不匹配");
        assertEquals(0, report.getSuccessRows(), "fail-fast 模式下不应有成功行");
    }

    /**
     * 测试 readWithReport 返回 JSON 友好的结果
     */
    @Test
    void importReportSuccessTest() throws Exception {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);

        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        
        cn.creekmoon.excel.core.R.report.ImportOptions options = new cn.creekmoon.excel.core.R.report.ImportOptions();
        options.setCollectSuccessData(false); // 不收集成功数据（节省内存）
        options.setCollectRowErrors(true);
        
        cn.creekmoon.excel.core.R.report.ImportReport<Student> report = excelImport
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .readWithReport(options);
        
        // 验证报告内容
        assertEquals(1000, report.getSuccessRows(), "成功行数应为1000");
        assertEquals(0, report.getErrorRows(), "错误行数应为0");
        assertEquals(1000, report.getTotalRows(), "总行数应为1000");
        assertNull(report.getData(), "未配置收集成功数据，应为null");
        assertNull(report.getGlobalErrorCode(), "无全局错误");
    }

    /**
     * 测试错误阈值停止
     */
    @Test
    void importReportErrorLimitTest() throws Exception {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);

        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        
        cn.creekmoon.excel.core.R.report.ImportOptions options = new cn.creekmoon.excel.core.R.report.ImportOptions();
        options.setMaxErrorRows(5); // 最多允许5个错误
        options.setCollectRowErrors(true);
        
        cn.creekmoon.excel.core.R.report.ImportReport<Student> report = excelImport
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge) // 这个会因为某些行数据问题失败
                .addConvertAndMustExist("必填字段不存在", Student::setEmail) // 故意配置不存在的必填项，制造错误
                .readWithReport(options);
        
        // 验证错误阈值触发
        assertTrue(report.getErrorRows() >= 5, "错误数应达到或超过阈值");
        assertEquals("ERROR_LIMIT_REACHED", report.getGlobalErrorCode(), "应该触发错误阈值");
        assertTrue(report.getTotalRows() < 1000, "应该提前终止，不读完全部");
    }

    /**
     * 测试收集成功数据的场景
     */
    @Test
    void importReportWithSuccessDataTest() throws Exception {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);

        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        
        cn.creekmoon.excel.core.R.report.ImportOptions options = new cn.creekmoon.excel.core.R.report.ImportOptions();
        options.setCollectSuccessData(true); // 收集成功数据
        options.setMaxSuccessData(100); // 最多收集100条
        
        cn.creekmoon.excel.core.R.report.ImportReport<Student> report = excelImport
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .readWithReport(options);
        
        // 验证成功数据收集
        assertNotNull(report.getData(), "应该收集成功数据");
        assertEquals(100, report.getData().size(), "应该只收集前100条");
        assertEquals(1000, report.getSuccessRows(), "成功行数应为1000");
        
        // 验证数据内容
        Student firstStudent = report.getData().get(0);
        assertNotNull(firstStudent.getUserName(), "数据应该正确填充");
    }

    /**
     * 测试 getAll 和 getReport 混合使用
     */
    @Test
    void importMixedGetAllAndReportTest() throws Exception {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);

        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        
        // 使用传统方式 getAll，但仍然可以获取 report
        TitleReader<Student> reader = excelImport
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime);
        
        // 使用 getAll 获取成功数据
        List<Student> students = reader.getAll();
        
        // 同时可以获取 report 查看详细统计
        cn.creekmoon.excel.core.R.report.ImportReport<Student> report = reader.getReport();
        
        // 验证两者一致性
        assertEquals(students.size(), report.getSuccessRows(), "getAll 返回数量应与 report 成功数一致");
        assertEquals(1000, students.size(), "应该成功读取1000条");
        assertEquals(1000, report.getTotalRows(), "总行数应为1000");
        assertEquals(0, report.getErrorRows(), "错误行数应为0");
        
        // report 的 data 字段应为空（因为没有配置 collectSuccessData）
        assertNull(report.getData(), "未配置收集，data 应为 null");
    }
}