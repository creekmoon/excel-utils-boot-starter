package cn.creekmoon.excelUtils.example;

import cn.creekmoon.excelUtils.converter.DateConverter;
import cn.creekmoon.excelUtils.converter.IntegerConverter;
import cn.creekmoon.excelUtils.converter.LocalDateTimeConverter;
import cn.creekmoon.excelUtils.core.AsyncTaskState;
import cn.creekmoon.excelUtils.core.ExcelExport;
import cn.creekmoon.excelUtils.core.ExcelImport;
import cn.creekmoon.excelUtils.core.SheetReader;
import cn.creekmoon.excelUtils.example.config.exception.MyNewException;
import cn.hutool.core.util.RandomUtil;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

@Api(tags = "测试API")
@RestController
@Slf4j
public class ExampleController {

    // key=taskId  value=异步状态  这里模拟保存到redis中
    private static final Map<String, AsyncTaskState> taskId2TaskState = new ConcurrentHashMap<>();

    @GetMapping(value = "/exportExcel")
    @ApiOperation("导出")
    public void exportExcel(Integer size, HttpServletRequest request, HttpServletResponse response) throws IOException {
        /*查询数据*/
        ArrayList<Student> result = createStudentList(size != null ? size : 60_000);
        ArrayList<Student> result2 = createStudentList(size != null ? size : 10_000);

        ExcelExport.create(Student.class)
                /*第一个标签页*/
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(result)
                /*第二个标签页*/
                .switchSheet(Student.class)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(result2)
                .response(response);
    }

    @GetMapping(value = "/exportExcelByStyle")
    @ApiOperation("导出(并设置style)")
    public void exportExcel5(HttpServletRequest request, HttpServletResponse response) throws IOException {
        ArrayList<Student> result = createStudentList(60_000);
        ArrayList<Teacher> result2 = createTeacherList(60_000);

        ExcelExport excelExport = ExcelExport.create();

        /*定义一个全局的数据样式  double是千分号和保留两位小数  int是千分号,保留整数*/
        short dataFormat_double = excelExport.getBigExcelWriter().getWorkbook().createDataFormat().getFormat("#,##0.00");
        short dataFormat_int = excelExport.getBigExcelWriter().getWorkbook().createDataFormat().getFormat("#,##0");

        excelExport
                .switchSheet("first_sheet", Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .setDataStyle(cellStyle ->
                {
                    cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                })
                .addTitle("年龄", Student::getAge)
                .setDataStyle(student -> student.getAge() > 20,
                        cellStyle ->
                        {
                            cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
                            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        }
                )
                .setDataStyle(student -> student.getAge() > 20,
                        cellStyle ->
                        {
                            cellStyle.setDataFormat(dataFormat_double);
                        }
                )
                .write(result)
                .response(response);
    }

    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcel")
    @ApiOperation("导入数据")
    public void importExcelBySax(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws IOException {
        //判断这个方法的执行时间
        long start = System.currentTimeMillis();
        ExcelImport excelImport = ExcelImport.create(file)
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .read(student -> {
                    if (student.age == 76) {
                        System.out.println("[年龄==76标记为错误]" + student);
                        throw new MyNewException("年龄==76");
                    }
                    System.out.println("[正常读取到对象]" + student);
                })
                .response(response);
        System.out.println("错误次数:" + excelImport.getErrorCount());
        //判断这个方法的执行时间
        long end = System.currentTimeMillis();
        System.out.println("执行时间:" + (end - start));
    }


    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcelByCell")
    @ApiOperation("导入数据(读取指定单元格)")
    public void importExcelByCell(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws IOException {

        ExcelImport excelImport = ExcelImport.create(file);
        SheetReader<Student> studentSheetReader = excelImport
                .switchSheet(0, Student::new)
                //读取指定单元格
                .addSingleCellReader("B1", x -> System.out.println("读取到B1单元格:" + x))
                .addSingleCellReader("D1", x -> System.out.println("读取到D1单元格:" + x))
                .addSingleCellReader("C5", x -> System.out.println("读取到C5单元格:" + x))
                .addSingleCellReader("E6", x -> System.out.println("读取到E6单元格:" + x))
                //从第二行开始,正常读取列
                .indexConfig(1)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime);


        long count = studentSheetReader
                .readAll()
                .stream()
                .peek(student -> studentSheetReader.setResult(student, "读取完毕"))
                .count();

        System.out.println("读取到students数量:" + count);
        studentSheetReader.response(response);
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

    private Teacher createNewTeacher() {
        Teacher teacher = new Teacher();
        //随机年龄
        teacher.setWorkYear(RandomUtil.randomInt(1, 10));
        teacher.setTeacherName(RandomUtil.randomString(5));
        return teacher;
    }

    private ArrayList<Student> createStudentList(int size) {
        ArrayList<Student> result = new ArrayList<>();
        //加入数据
        for (int i = 0; i < size; i++) {
            result.add(createNewStudent());
        }
        return result;
    }

    private ArrayList<Teacher> createTeacherList(int size) {
        ArrayList<Teacher> result = new ArrayList<>();
        //加入数据
        for (int i = 0; i < size; i++) {
            result.add(createNewTeacher());
        }
        return result;
    }


}
