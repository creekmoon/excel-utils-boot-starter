package cn.creekmoon.excel.example;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.core.W.title.TitleWriter;
import cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle;
import cn.hutool.core.util.RandomUtil;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.CompletableFuture;

import static cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle.of;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.*;
import static cn.creekmoon.excel.example.Student.createStudentList;
import static cn.creekmoon.excel.example.Teacher.createTeacherList;

@Tag(name = "测试API")
@RestController
@Slf4j
public class ExampleController {


    @GetMapping(value = "/exportExcel")
    @Operation(summary = "导出")
    public void exportExcel(Integer size, HttpServletRequest request, HttpServletResponse response) throws IOException {
        /*查询数据*/
        ArrayList<Student> result = createStudentList(size != null ? size : 60_000);
        ArrayList<Student> result2 = createStudentList(size != null ? size : 10_000);
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

        excelExport.response(response);
    }

    @GetMapping(value = "/exportExcelByStyle")
    @Operation(summary = "导出(并设置style)")
    public void exportExcel5(HttpServletRequest request, HttpServletResponse response) throws IOException {
        ArrayList<Student> result = createStudentList(60_000);
        ArrayList<Teacher> result2 = createTeacherList(60_000);

        /*定义一个全局的数据样式  保留两位小数 */
        ExcelCellStyle customCellStyle = new ExcelCellStyle((workbook, style) -> style.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00")));


        ExcelExport excelExport = ExcelExport.create();
        TitleWriter<Student> studentTitleWriter = excelExport.switchNewSheet(Student.class);
        studentTitleWriter
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName,
                        of(LIGHT_GREEN, x -> x.getAge() > 50))
                .addTitle("年龄", Student::getAge,
                        of(customCellStyle, student -> student.getAge() > 50)
                )
                .write(result);
        excelExport.response(response);
    }

    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcel")
    @Operation(summary = "导入数据")
    public void importExcelBySax(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws Exception {
        if (file==null) {
            file = new MockMultipartFile(
                    "导入测试.xlsx",
                    "导入测试.xlsx",
                    "application/vnd.ms-excel",
                    new org.springframework.core.io.ClassPathResource("导入测试.xlsx").getInputStream()
            );
        }
        /*读取导入文件*/
        ExcelImport excelImport = ExcelImport.create(file);
        TitleReaderResult<Student> sheet1 = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .read(x -> log.info(x.toString()));

        TitleReaderResult<Student> sheet2 = excelImport.switchSheet(1, Student::new)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("邮箱", Student::setEmail)
                .read(x -> {
                    if ("lisi@qq.com".equals(x.getEmail())) {
                        Thread.sleep(2000);
                        throw new RuntimeException("错误!");
                    }
                    log.info(x.toString());
                });

        excelImport.response(response);
    }


    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcelByCell")
    @Operation(summary = "导入数据(读取指定单元格)")
    public void importExcelByCell(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws Exception {
        /*从资源文件中读取测试数据*/
        if (file==null) {
             file = new MockMultipartFile(
                    "导入测试(读取指定单元格).xlsx",
                    "导入测试(读取指定单元格).xlsx",
                    "application/vnd.ms-excel",
                    new org.springframework.core.io.ClassPathResource("导入测试(读取指定单元格).xlsx").getInputStream()
            );
        }



        ExcelImport excelImport = ExcelImport.create(file);
        ReaderResult<Student> readerResult = excelImport.switchSheetAndUseCellReader(1, Student::new)
                .addConvertAndMustExist("A1", Student::setUserName)
                .read(x -> {
                    log.info(x.toString());
                });

        System.out.println("readerResult.getErrorReport() = " + readerResult.errorReport);
        excelImport.response(response);

    }

    /**
     * 导出Excel并使用range方法控制写入范围
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @GetMapping(value = "/exportExcelWithRange")
    @Operation(summary = "导出(使用range方法控制写入范围)")
    public void exportExcelWithRange(HttpServletRequest request, HttpServletResponse response) throws IOException {
        /*创建测试数据*/
        ArrayList<Student> studentList = createStudentList(100);

        ExcelExport excelExport = ExcelExport.create();

        /*示例1: 从第5行开始写入标题和数据，数据写入到第20行*/
        excelExport.switchNewSheet("从第5行开始", Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .range(5, 19)  // 标题在第5行(索引4)，数据写入到第20行(索引19)
                .write(studentList);

        /*示例2: 从第10行开始写入，不限制结束行*/
        excelExport.switchNewSheet("从第10行开始", Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge,
                        of(LIGHT_GREEN, x -> x.getAge() > 50))
                .addTitle("邮箱", Student::getEmail)
                .range(9)  // 标题在第10行(索引9)，数据从第11行开始写到结束
                .write(studentList);

        /*示例3: 完整指定标题行、首行数据和末行数据*/
        excelExport.switchNewSheet("完整指定范围", Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName,
                        of(LIGHT_ORANGE, x -> true))
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .range(2, 6, 15)  // 标题在第3行(索引2)，数据从第5行(索引4)开始，写到第16行(索引15)
                .write(studentList);


        excelExport.response(response);
    }

    /**
     * 导入Excel并使用range方法控制读取范围
     *
     * @param file
     * @param request
     * @param response
     * @throws Exception
     */
    @PostMapping(value = "/importExcelWithRange")
    @Operation(summary = "导入数据(使用range方法控制读取范围)")
    public void importExcelWithRange(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws Exception {
        if (file==null) {
            file = new MockMultipartFile(
                    "导入测试.xlsx",
                    "导入测试.xlsx",
                    "application/vnd.ms-excel",
                    new org.springframework.core.io.ClassPathResource("导入测试.xlsx").getInputStream()
            );
        }

        /*读取导入文件*/
        ExcelImport excelImport = ExcelImport.create(file);

        /*示例1: 只读取前20行数据*/
        TitleReaderResult<Student> sheet1 = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .range(0, 19)  // 标题在第1行(索引0)，读取数据到第20行(索引19)
                .read(x -> log.info("Sheet1数据: {}", x.toString()));

        /*示例2: 从第5行开始读取标题和数据*/
        TitleReaderResult<Student> sheet2 = excelImport.switchSheet(1, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .range(4)  // 标题在第5行(索引4)，数据从第6行开始读取
                .read(x -> log.info("Sheet2数据: {}", x.toString()));

        excelImport.response(response);
    }

    /**
     * 导入Excel并读取同一个sheet中的多个不同类型的表格
     * 读取的是通过 /exportExcelWithReset 接口生成的 Excel 文件
     *
     * @param response
     * @throws Exception
     */
    @GetMapping(value = "/importExcelWithMultipleTables")
    @Operation(summary = "导入数据(读取同一sheet中多个表格)")
    public void importExcelWithMultipleTables(HttpServletResponse response) throws Exception {
        /*从资源文件中读取测试数据*/
        MultipartFile file = new MockMultipartFile(
                "导入测试Sheet页存在多个表格.xlsx",
                "导入测试Sheet页存在多个表格.xlsx",
                "application/vnd.ms-excel",
                new org.springframework.core.io.ClassPathResource("导入测试Sheet页存在多个表格.xlsx").getInputStream()
        );

        ExcelImport excelImport = ExcelImport.create(file,true);
        /*读取第一个表格：Student（第0行开始）*/
        log.info("========== 开始读取 Student 表格 ==========");
        TitleReader<Student> reader = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .range(0, 50);
        reader.read(student -> log.info("读取Student: 用户名={}, 全名={}, 年龄={}, 邮箱={}", 
                        student.getUserName(), student.getFullName(), student.getAge(), student.getEmail()));

        /*使用 reset 方法读取第二个表格：Teacher（约在第52行）*/
        log.info("========== 开始读取 Teacher 表格 ==========");
        TitleReader<Teacher> teacherReader = reader.reset(Teacher::new)
                .addConvert("教师姓名", Teacher::setTeacherName)
                .addConvert("工作年限", IntegerConverter::parse, Teacher::setWorkYear)
                .range(52, 82);
        teacherReader.read(teacher -> log.info("读取Teacher: 教师姓名={}, 工作年限={}", 
                        teacher.getTeacherName(), teacher.getWorkYear()));

        /*再次使用 reset 读取第三个表格：Student（约在第84行）*/
        log.info("========== 开始读取第二次 Student 表格 ==========");
        teacherReader.reset(Student::new)
                .addConvert("基本信息::用户名", Student::setUserName)
                .addConvert("基本信息::全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .range(84)
                .read(student -> log.info("读取Student(多级表头): 用户名={}, 全名={}, 年龄={}", 
                        student.getUserName(), student.getFullName(), student.getAge()));

        /*
         * 注意：此示例使用了 reset 方法读取多个不同类型的数据
         * reset 创建的 Reader 不会参与导入结果文件的生成
         * 因此这里不调用 excelImport.response(response)
         * 
         * 如果需要生成包含所有数据验证结果的文件，建议：
         * 1. 使用多个 ExcelImport 实例分别处理
         * 2. 或者将不同类型的数据放在不同的 Sheet 页中
         */
        // excelImport.response(response);  // 不调用，因为结果不完整
        
        log.info("========== 数据读取完成 ==========");
    }

    /**
     * 导出Excel并使用reset方法在同一个sheet中写入多个不同类型的表格
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @GetMapping(value = "/exportExcelWithReset")
    @Operation(summary = "导出(使用reset方法在同一sheet中写入多个表格)")
    public void exportExcelWithReset(HttpServletRequest request, HttpServletResponse response) throws IOException {
        /*创建测试数据*/
        ArrayList<Student> studentList = createStudentList(50);
        ArrayList<Teacher> teacherList = createTeacherList(30);

        ExcelExport excelExport = ExcelExport.create();

        /*在同一个sheet中写入不同类型的多个表格*/
        TitleWriter<Student> writer = excelExport.switchNewSheet("多表格Sheet", Student.class);
        
        /*写入第一个表格：Student*/
        writer.addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .write(studentList);

        /*使用reset方法重置writer，在同一个sheet中写入第二个表格：Teacher*/
        writer.reset(Teacher.class)
                .addTitle("教师姓名", Teacher::getTeacherName)
                .addTitle("工作年限", Teacher::getWorkYear,
                        of(LIGHT_GREEN, x -> x.getWorkYear() > 5))
                .write(teacherList);

        /*可以继续reset，写入第三个表格（这里再次写入Student作为示例）*/
        writer.reset(Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName,
                        of(LIGHT_ORANGE, x -> true))
                .addTitle("年龄", Student::getAge)
                .range(85)  // 也可以手动指定起始位置
                .write(studentList.subList(0, 10));

        excelExport.response(response);
    }


}
