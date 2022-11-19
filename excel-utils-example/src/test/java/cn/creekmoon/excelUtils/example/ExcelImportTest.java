package cn.creekmoon.excelUtils.example;

import cn.creekmoon.excelUtils.converter.DateConverter;
import cn.creekmoon.excelUtils.converter.IntegerConverter;
import cn.creekmoon.excelUtils.converter.LocalDateTimeConverter;
import cn.creekmoon.excelUtils.core.ExcelExport;
import cn.creekmoon.excelUtils.core.ExcelImport;
import cn.creekmoon.excelUtils.core.PathFinder;
import cn.creekmoon.excelUtils.threadPool.CleanTempFilesExecutor;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.RandomUtil;
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

import java.io.BufferedInputStream;
import java.net.URI;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
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


    @Test
    public void exportTest() throws Exception {
        AtomicInteger count = new AtomicInteger(0);
        int targetCount = 600;
        //构造请求
        ResultActions resultActions = this.mvc.perform(MockMvcRequestBuilders
                .get(URI.create("/exportExcel"))
                .queryParam("size", String.valueOf(targetCount)));
        resultActions.andReturn().getResponse().setCharacterEncoding("UTF-8");
        resultActions
                .andExpect(MockMvcResultMatchers.status().isOk())
                .andDo(x -> {
                    MockMultipartFile mockFile = new MockMultipartFile("mockFile", x.getResponse().getContentAsByteArray());
                    ExcelImport<Student> read = ExcelImport.create(mockFile, Student::new)
                            .addConvert("用户名", Student::setUserName)
                            .addConvert("全名", Student::setFullName)
                            .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                            .addConvert("邮箱", Student::setEmail)
                            .addConvert("生日", DateConverter::parse, Student::setBirthday)
                            .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                            .read(data -> {
                                count.incrementAndGet();
                            });
                });

        // 断言返回值是600行
        Assertions.assertEquals(targetCount, count.get());
    }


    @Test
    public void importTest() throws Exception {
        //构建一个Excel 并获取文件对象
        List<Student> students = Arrays.asList(
                Student.builder().fullName("first_full_name").age(1).build(),
                Student.builder().fullName("second_full_name").age(2).build());
        String taskId = ExcelExport.create("test", Student.class)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("生日", Student::getBirthday)
                .write(students)
                .stopWrite();
        Assertions.assertNotNull(taskId);
        BufferedInputStream inputStream = FileUtil.getInputStream(PathFinder.getAbsoluteFilePath(taskId));
        Assertions.assertNotNull(inputStream);

        //读取excel
        MockMultipartFile testExcel = new MockMultipartFile("test", inputStream);
        List<Student> result = ExcelImport.create(testExcel, Student::new)
                .addConvertAndMustExist("全名", Student::setFullName)
                .addConvertAndSkipEmpty("年龄", IntegerConverter::parse, Student::setAge)
                .addConvertAndSkipEmpty("生日", DateConverter::parse, Student::setBirthday)
                .readAll(ExcelImport.ConvertStrategy.SKIP_ALL_IF_FAIL);

        Assertions.assertEquals(result.size(), 2);
        CleanTempFilesExecutor.cleanTeamFileNow(taskId);
        Assertions.assertFalse(FileUtil.exist(PathFinder.getAbsoluteFilePath(taskId)));

    }

    private Student createNewStudent() {
        Student student = new Student();
        //随机年龄
        student.setAge(RandomUtil.randomInt(1, 100));
        student.setBirthday(new Date());
        //随机生成邮箱
        student.setEmail(RandomUtil.randomString(10) + "@qq.com");
        //随机生成时间
        student.setExpTime(LocalDateTime.now());
        student.setFullName(RandomUtil.randomString(5));
        student.setUserName(RandomUtil.randomString(5));
        student.setBirthday(new Date());
        return student;
    }

    private ArrayList<Student> createStudentList(int size) {
        ArrayList<Student> result = new ArrayList<>();
        //加入数据 六十万
        for (int i = 0; i < size; i++) {
            result.add(createNewStudent());
        }
        return result;
    }

}