package cn.creekmoon.excelUtils.example;

import cn.creekmoon.excelUtils.core.ExcelExport;
import cn.hutool.core.util.RandomUtil;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;

@Tag(name = "测试API")
@RestController
@Slf4j
public class TestController {


    @GetMapping(value = "/exampleTest")
    @Operation(summary = "导出")
    public void exportExcel(Integer size, HttpServletRequest request, HttpServletResponse response) throws IOException {
        /*查询数据*/
        ExcelExport.create(Student.class)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .write(createStudentList(size != null ? size : 60_000))
                .response(response);
    }


    /**
     * {
     * "userName": "johndoe",
     * "fullName": "John Doe",
     * "email": "johndoe@example.com",
     * "age": 25,
     * "birthday": "1998-05-12",
     * "expTime": "2023-06-04 16:42:26"
     * }
     *
     * @return
     */
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


    /**
     * 模拟数据
     * <p>
     * {
     * "teacherName": "张三",
     * "workYear": 5
     * }
     *
     * @return
     */
    private Teacher createNewTeacher() {
        Teacher teacher = new Teacher();
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
