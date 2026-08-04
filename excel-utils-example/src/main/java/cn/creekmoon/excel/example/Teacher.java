package cn.creekmoon.excel.example;

import lombok.*;

import java.util.ArrayList;
import java.util.concurrent.ThreadLocalRandom;


/**
 * 注意导出的对象不能重写 equal和hashCode方法
 */
@Setter
@Getter
@Builder
@ToString
@NoArgsConstructor
@AllArgsConstructor
public class Teacher {


    String teacherName;
    Integer workYear;
    public static  Teacher createNewTeacher() {
        Teacher teacher = new Teacher();
        //随机年龄
        teacher.setWorkYear(ThreadLocalRandom.current().nextInt(1, 10));
        teacher.setTeacherName(randomString(5));
        return teacher;
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


    public static  ArrayList<Teacher> createTeacherList(int size) {
        ArrayList<Teacher> result = new ArrayList<>();
        //加入数据
        for (int i = 0; i < size; i++) {
            result.add(createNewTeacher());
        }
        return result;
    }
}
