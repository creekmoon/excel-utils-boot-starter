package cn.creekmoon.excelUtils.example;

import cn.creekmoon.excelUtils.converter.DateConverter;
import cn.creekmoon.excelUtils.converter.IntegerConverter;
import cn.creekmoon.excelUtils.converter.LocalDateTimeConverter;
import cn.creekmoon.excelUtils.core.*;
import cn.creekmoon.excelUtils.threadPool.CleanTempFilesExecutor;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ResourceUtil;
import cn.hutool.core.util.ReflectUtil;
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
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

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


    @Test
    public void exportTest() throws Exception {
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


    @Test
    public void importTest() throws Exception {
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

        /*检查导出结果是否按照预期生成*/
        ExcelExport excelExport = (ExcelExport) ReflectUtil.getFieldValue(read, "excelExport");
        Assertions.assertFalse(FileUtil.exist(PathFinder.getAbsoluteFilePath(excelExport.taskId)));
        read.generateResultFile();
        Assertions.assertTrue(FileUtil.exist(PathFinder.getAbsoluteFilePath(excelExport.taskId)));
        CleanTempFilesExecutor.cleanTempFileNow(excelExport.taskId);
        Assertions.assertFalse(FileUtil.exist(PathFinder.getAbsoluteFilePath(excelExport.taskId)));

    }

    @Test
    public void importTest2() throws Exception {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-cell-read.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);


        /*导入测试*/
        AtomicReference<String> B1_RESULT = new AtomicReference<String>(null);
        AtomicReference<String> D1_RESULT = new AtomicReference<String>(null);
        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        SheetReader<Student> studentSheetReader = excelImport
                .switchSheet(0, Student::new)
                .addSingleCellReader("B1", B1_RESULT::set)
                .addSingleCellReader("D1", D1_RESULT::set)
                .indexConfig(1)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime);
        List<Student> students = studentSheetReader.readAll();


        /*检查导出结果是否按照预期生成*/
        Assertions.assertEquals(students.size(), 4);
        Assertions.assertEquals(excelImport.getErrorCount().get(), 0);
        Assertions.assertEquals(B1_RESULT.get(), "李二狗");
        Assertions.assertEquals(D1_RESULT.get(), "李云龙");

    }
}