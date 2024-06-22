package cn.creekmoon.excel.example;


import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.core.W.title.TitleWriter;
import cn.hutool.core.io.FileUtil;
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

import java.net.URI;
import java.util.ArrayList;
import java.util.concurrent.atomic.AtomicInteger;

import static cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle.of;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.LIGHT_ORANGE;
import static cn.creekmoon.excel.example.Student.createStudentList;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

@AutoConfigureMockMvc
@SpringBootTest
@Slf4j
public class ExcelExportTest {

    @Autowired
    private MockMvc mvc;
    private MockHttpSession session;

    @BeforeEach
    public void init() {
        session = new MockHttpSession();
    }

    /**
     * 测试导出方法1
     *
     * @throws Exception
     */
    @Test
    public void exportTest1() throws Exception {

        /*查询数据*/
        ArrayList<Student> result = createStudentList( 60_000);
        ArrayList<Student> result2 = createStudentList(10_000);
        ExcelExport excelExport = ExcelExport.create();
        TitleWriter<Student> sheet0 = excelExport.switchNewSheet(Student.class);
        /*第一个标签页*/
        sheet0.addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName,
                        of(LIGHT_ORANGE, x -> true))
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(result);

        /*第二个标签页*/
        excelExport.switchNewSheet(Student.class)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(result2);

        String filePath = excelExport.stopWrite();
        System.out.println("filePath = " + filePath);
        Assertions.assertTrue(FileUtil.file(filePath).exists());
    }



}
