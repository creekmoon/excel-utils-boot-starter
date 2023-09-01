package cn.creekmoon.excelUtils.example;

import cn.creekmoon.excelUtils.converter.DateConverter;
import cn.creekmoon.excelUtils.converter.IntegerConverter;
import cn.creekmoon.excelUtils.converter.LocalDateTimeConverter;
import cn.creekmoon.excelUtils.core.*;
import cn.creekmoon.excelUtils.threadPool.CleanTempFilesExecutor;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ResourceUtil;
import cn.hutool.core.util.ReflectUtil;
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

import java.io.InputStream;
import java.net.URI;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

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
                    ExcelImport.create(mockFile, Student::new)
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
        ExcelImport read = ExcelImport.create(mockMultipartFile)
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .read(false, student -> {
                    log.info("[测试导入] object={}", student);
                    count.getAndIncrement();
                });
        Assertions.assertEquals(1000, count.get());

        /*第二个sheet导入测试*/
        List<Student> students = read.switchSheet(1, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .readAll();
        Assertions.assertEquals(4, students.size());
        for (Student student : students) {
            read.setResult(student, ExcelConstants.IMPORT_SUCCESS_MSG);
        }

        //检查临时文件是否能够正常生成和清理
        ExcelExport excelExport = (ExcelExport) ReflectUtil.getFieldValue(read, "excelExport");
        Assertions.assertFalse(FileUtil.exist(PathFinder.getAbsoluteFilePath(excelExport.taskId)));
        read.generateResultFile();
        Assertions.assertTrue(FileUtil.exist(PathFinder.getAbsoluteFilePath(excelExport.taskId)));
        CleanTempFilesExecutor.cleanTempFileNow(excelExport.taskId);
        Assertions.assertFalse(FileUtil.exist(PathFinder.getAbsoluteFilePath(excelExport.taskId)));

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
        SheetReader<Student> studentSheetReader = excelImport
                .switchSheet(0, Student::new)
                .range(0, 3, 4)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime);
        List<Student> students = studentSheetReader.readAll();


        /*检查是否能够正确读取*/
        Assertions.assertEquals(students.size(), 2);
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
        Student sheet1 = excelImport
                .switchSheetAndUseCellReader(0, Student::new)
                .addConvert("B1", Student::setUserName)
                .addConvert("D1", Student::setFullName)
                .addConvert(0, 5, IntegerConverter::parse, Student::setAge)
                .read();
        Assertions.assertEquals(sheet1.getUserName(), "李二狗");
        Assertions.assertEquals(sheet1.getFullName(), "李云龙");
        Assertions.assertEquals(sheet1.getAge(), 2000);


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
}