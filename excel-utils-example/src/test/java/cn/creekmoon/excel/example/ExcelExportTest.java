package cn.creekmoon.excel.example;


import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.core.W.title.TitleWriter;
import cn.hutool.poi.excel.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.mock.web.MockHttpSession;
import org.springframework.test.web.servlet.MockMvc;

import java.io.File;
import java.util.Collections;
import java.util.ArrayList;

import static cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle.of;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.LIGHT_ORANGE;
import static cn.creekmoon.excel.example.Student.createStudentList;

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
        ArrayList<Student> result = createStudentList(60_000);
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

        File file = excelExport.stopWrite();
        System.out.println("filePath = " + file.getAbsolutePath());
        Assertions.assertTrue(file.exists());

        //验证行数
        Assertions.assertEquals(60_000 + 2, ExcelUtil.getReader(file, 0).getRowCount());
        Assertions.assertEquals(10_000 + 1, ExcelUtil.getReader(file, 1).getRowCount());


    }

    @Test
    public void exportTest2() throws Exception {

        /*查询数据*/
        ArrayList<Student> result = createStudentList(1_000);
        ArrayList<Student> result2 = createStudentList(1_000);
        ExcelExport excelExport = ExcelExport.create();
        TitleWriter<Student> sheet0 = excelExport.switchNewSheet(Student.class);
        /*第一个标签页*/
        TitleWriter<Student> studentTitleWriter = sheet0.addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName,
                        of(LIGHT_ORANGE, x -> true))
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime);


        /*第二个标签页*/
        TitleWriter<Student> studentTitleWriter1 = excelExport.switchNewSheet(Student.class)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime);
        
        /*交替写*/
        studentTitleWriter.write(result);
        studentTitleWriter1.write(result2);
        studentTitleWriter.write(result);
        studentTitleWriter1.write(result2);


        File file = excelExport.stopWrite();
        System.out.println("filePath = " + file.getAbsolutePath());
        Assertions.assertTrue(file.exists());

        //验证行数
        Assertions.assertEquals(1_000*2 + 2, ExcelUtil.getReader(file, 0).getRowCount());
        Assertions.assertEquals(1_000*2 + 1, ExcelUtil.getReader(file, 1).getRowCount());


    }

    @Test
    public void exportEmptyThenAppendTest() throws Exception {
        ArrayList<Student> result = createStudentList(10);
        ExcelExport excelExport = ExcelExport.create();

        TitleWriter<Student> writer = excelExport.switchNewSheet(Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName,
                        of(LIGHT_ORANGE, x -> true))
                .addTitle("年龄", Student::getAge);

        // 首次空写：只写表头
        writer.write(Collections.emptyList());
        // 追加写：应从表头之后开始写数据
        writer.write(result);
        writer.write(result);

        File file = excelExport.stopWrite();
        Assertions.assertTrue(file.exists());
        // 多级表头深度=2，所以表头占2行；数据写入两次，共20行
        Assertions.assertEquals(20 + 2, ExcelUtil.getReader(file, 0).getRowCount());
    }
}
