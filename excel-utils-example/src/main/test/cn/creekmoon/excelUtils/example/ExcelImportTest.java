package cn.creekmoon.excelUtils.example;

import cn.creekmoon.excelUtils.converter.DateConverter;
import cn.creekmoon.excelUtils.converter.IntegerConverter;
import cn.creekmoon.excelUtils.converter.LocalDateTimeConverter;
import cn.creekmoon.excelUtils.core.ExcelImport;
import cn.hutool.core.util.RandomUtil;
import lombok.extern.slf4j.Slf4j;
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

import java.net.URI;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;

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
    public void unitTest() throws Exception {

        ResultActions resultActions = this.mvc.perform(MockMvcRequestBuilders.get(URI.create("/exportExcel")));
        // 解决中文乱码问题
        resultActions.andReturn().getResponse().setCharacterEncoding("UTF-8");
        resultActions
                .andExpect(MockMvcResultMatchers.status().isOk())
                // .andExpect(MockMvcResultMatchers.content().string(containsString("OpenId B")))
                .andDo(x -> {
                    MockMultipartFile mockFile = new MockMultipartFile("mockFile", x.getResponse().getContentAsByteArray());
                    System.out.println(mockFile);
                    ExcelImport<Student> read = ExcelImport.create(mockFile, Student::new)
                            .addConvert("用户名", Student::setUserName)
                            .addConvert("全名", Student::setFullName)
                            .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                            .addConvert("邮箱", Student::setEmail)
                            .addConvert("生日", DateConverter::parse, Student::setBirthday)
                            .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                            .read(student -> {
                                System.out.println(student);
                            });
                });
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