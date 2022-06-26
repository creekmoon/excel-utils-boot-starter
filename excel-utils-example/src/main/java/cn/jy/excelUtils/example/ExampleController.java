package cn.jy.excelUtils.example;

import cn.hutool.core.date.DatePattern;
import cn.hutool.core.util.RandomUtil;
import cn.jy.excelUtils.converter.DateConverter;
import cn.jy.excelUtils.converter.IntegerConverter;
import cn.jy.excelUtils.converter.LocalDateTimeConverter;
import cn.jy.excelUtils.core.AsyncTaskState;
import cn.jy.excelUtils.core.ExcelExport;
import cn.jy.excelUtils.core.ExcelImport;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;

@RestController("/test")
public class ExampleController {

    // key=taskId  value=异步状态  这里模拟保存到redis中
    private static final Map<String, AsyncTaskState> taskId2TaskState = new ConcurrentHashMap<>();

    @GetMapping(value = "/exportExcel")
    @ApiOperation("单次查询,并导出数据")
    public void exportExcel(HttpServletRequest request, HttpServletResponse response) throws IOException {
        ArrayList<Student> result = createStudentList(60_000);
        ExcelExport.create("lalala", Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(result)
                .response(response);
    }


    @GetMapping(value = "/exportExcel2")
    @ApiOperation("多次查询,并导出数据")
    public void exportExcel2(HttpServletRequest request, HttpServletResponse response) throws IOException {

        /*构建表头*/
        ExcelExport<Student> excelExport = ExcelExport.create("lalala", Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime);
        //模拟查询
        for (int i = 0; i < 3; i++) {
            excelExport.write(createStudentList(25_000));
        }
        /*返回数据*/
        excelExport.response(response);
    }


    @GetMapping(value = "/exportExcel3")
    @ApiOperation("构建多级表头,导出数据")
    public void exportExcel3(HttpServletRequest request, HttpServletResponse response) throws IOException {
        ArrayList<Student> result = createStudentList(60_000);
        ExcelExport.create("lalala", Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName)
                .addTitle("额外附加信息::年龄", Student::getAge)
                .addTitle("额外附加信息::邮箱", Student::getEmail)
                .addTitle("额外附加信息::系统数据::生日", Student::getBirthday)
                .addTitle("额外附加信息::系统数据::过期时间", Student::getExpTime)
                .write(result)
                .response(response);
    }

    @GetMapping(value = "/exportExcel4")
    @ApiOperation("异步导出数据, 会返回的taskId")
    public AsyncTaskState exportExcel4(HttpServletRequest request, HttpServletResponse response) throws IOException {
        /*尝试次数*/
        AtomicInteger tryCount = new AtomicInteger(2000);
        /*开始导入*/
        AsyncTaskState taskState = ExcelExport.create("lalala", Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName)
                .addTitle("额外附加信息::年龄", Student::getAge)
                .addTitle("额外附加信息::邮箱", Student::getEmail)
                .addTitle("额外附加信息::系统数据::生日", Student::getBirthday)
                .addTitle("额外附加信息::系统数据::过期时间", Student::getExpTime)
                //写入数据Size==0时 就会自动停止写入
                .writeAsync(() -> tryCount.decrementAndGet() < 0 ? new ArrayList<>() : createStudentList(500),
                        state -> {
                            System.out.println(state);
                            //保存到redis中
                            taskId2TaskState.put(state.getTaskId(), state);
                        });
        return taskState;
    }


    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcel1")
    public void importExcel(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws IOException {
        //判断这个方法的执行时间
        long start = System.currentTimeMillis();
        ExcelImport.create(file, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .read(student -> {
                    System.out.println(student);
                })
                .response(response);

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
    @PostMapping(value = "/importExcel2")
    public AsyncTaskState importExcel2(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws IOException {
        //判断这个方法的执行时间
        long start = System.currentTimeMillis();
        AsyncTaskState asyncTaskState = ExcelImport.create(file, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .readAsync(
                        student -> {
                            //System.out.println(student);
                        }, state -> {
                            System.out.println(state);
                        }
                );
        System.out.println(asyncTaskState);
        //判断这个方法的执行时间
        long end = System.currentTimeMillis();
        System.out.println("执行时间:" + (end - start));
        return asyncTaskState;
    }

    @GetMapping(value = "/followTaskId")
    @ApiOperation("跟踪异步状态")
    public AsyncTaskState followTaskId(String taskId) throws IOException {
        return taskId2TaskState.get(taskId);
    }

    @GetMapping(value = "/getResultByTaskId")
    @ApiOperation("获取异步结果")
    public void getResultByTaskId(String taskId, HttpServletRequest request, HttpServletResponse response) throws IOException {
        try {
            /*响应导出结果*/
            ExcelExport.response(taskId, "result[" + LocalDateTime.now().format(DateTimeFormatter.ofPattern(DatePattern.NORM_DATETIME_MINUTE_PATTERN)) + "].xlsx", response);
        } finally {
            /*清除临时文件*/
            ExcelExport.cleanTempFile(taskId);
        }
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
