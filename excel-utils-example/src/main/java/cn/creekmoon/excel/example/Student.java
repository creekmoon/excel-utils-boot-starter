package cn.creekmoon.excel.example;

import lombok.*;

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.ThreadLocalRandom;


/**
 * 注意导出的对象不能重写 equal和hashCode方法
 */
@Setter
@Getter
@Builder
@NoArgsConstructor
@AllArgsConstructor
@ToString
public class Student {


    String userName;
    String fullName;
    String email;
    Integer age;
    Date birthday;
    LocalDateTime expTime;


    public static Student createNewStudent() {
        Student student = new Student();
        //随机年龄
        student.setAge(ThreadLocalRandom.current().nextInt(1, 100));
        student.setBirthday(new Date());
        //随机生成邮箱
        student.setEmail(randomString(10) + "@qq.com");
        //随机生成时间
        student.setExpTime(LocalDateTime.now());
        student.setFullName(randomString(5));
        student.setUserName(randomString(5));
        student.setBirthday(new Date());
        return student;
    }


    /**
     * 生成指定长度的随机小写字母字符串(替代 Hutool RandomUtil.randomString)
     */
    private static String randomString(int length) {
        StringBuilder sb = new StringBuilder(length);
        for (int i = 0; i < length; i++) {
            sb.append((char) ('a' + ThreadLocalRandom.current().nextInt(26)));
        }
        return sb.toString();
    }


    public static  ArrayList<Student> createStudentList(int size) {
        ArrayList<Student> result = new ArrayList<>();
        //加入数据
        for (int i = 0; i < size; i++) {
            result.add(createNewStudent());
        }
        return result;
    }
}
