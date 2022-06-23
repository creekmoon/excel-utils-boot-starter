package cn.jy.excelUtils.example;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.RandomUtil;
import cn.jy.excelUtils.core.AsyncReadState;
import cn.jy.excelUtils.core.ExcelExport;
import cn.jy.excelUtils.core.ExcelImport;
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

@RestController("/test")
public class ExampleController {

    /**
     * 导出
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @GetMapping(value = "/exportExcel")
    public void exportExcel(HttpServletRequest request, HttpServletResponse response) throws IOException {
        ArrayList<Student> result = new ArrayList<>();
        //循环一万次加入数据
        for (int i = 0; i < 200000; i++) {
            result.add(getStudent());
        }
        ExcelExport.create("lalala", Student.class)
                .addTitle("姓名", Student::getUserName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("过期时间", Student::getExpTime)
                .write(result)
                .write(result)
                .write(result)
                .response(response);

    }

    private Student getStudent() {
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


    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcel")
    public void importExcel(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws IOException {
        //判断这个方法的执行时间
        long start = System.currentTimeMillis();

        ExcelImport.create(file, Student::new)
                .addConvert("姓名", Student::setUserName)
                .addConvert("年龄", Integer::valueOf, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("过期时间", x -> {
                    return DateUtil.parse(x.substring(0, 10));
                }, Student::setBirthday)
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
    @PostMapping(value = "/importExcelSax")
    public void importExcelSax(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws IOException {
        //判断这个方法的执行时间
        long start = System.currentTimeMillis();
        AsyncReadState asyncReadState = ExcelImport.create(file, Student::new)
                .addConvert("姓名", Student::setUserName)
                .addConvert("年龄", (String x) -> {
                    Integer integer = x.contains(".") ? Integer.valueOf(x.substring(0, x.indexOf("."))) : Integer.valueOf(x);
                    return integer;
                }, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("过期时间", x -> {
                    return DateUtil.parse(x.substring(0, 10));
                }, Student::setBirthday)
                .saxRead(
                        student -> {
                            //System.out.println(student);
                        }, state -> {
                            System.out.println(state);
                        }
                );
        System.out.println(asyncReadState);
        //判断这个方法的执行时间
        long end = System.currentTimeMillis();
        System.out.println("执行时间:" + (end - start));
    }
}
