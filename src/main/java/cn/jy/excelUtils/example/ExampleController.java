package cn.jy.excelUtils.example;

import cn.jy.excelUtils.core.ExcelExport;
import org.springframework.boot.web.server.WebServerFactoryCustomizer;
import org.springframework.boot.web.servlet.server.ConfigurableServletWebServerFactory;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;

@RestController
public class ExampleController  {


    @GetMapping(value = "/exportExcel")
    public void exportExcel( HttpServletRequest request, HttpServletResponse response) throws IOException {
        Student student = new Student();
        student.setAge(15);
        student.setBirthday(new Date());
        student.setEmail("lalal@hotmail.com");
        student.setExpTime(LocalDateTime.now());
        student.setFullName("lalala");
        student.setUserName("xixixi");
        student.setBirthday(new Date());
        ArrayList<Student> result = new ArrayList<>();
        result.add(student);
        result.add(student);
        result.add(student);
        result.add(student);
        ExcelExport.create("lalala",Student.class)
                .addTitle("学生::姓名",Student::getUserName)
                .addTitle("学生::年龄",Student::getAge)
                .addTitle("学生::邮箱",Student::getEmail)
                .addTitle("过期时间",Student::getExpTime)
                .write(result)
                .response(response);

    }


}
