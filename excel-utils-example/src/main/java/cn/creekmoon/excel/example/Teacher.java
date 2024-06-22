package cn.creekmoon.excel.example;

import cn.hutool.core.util.RandomUtil;
import lombok.*;

import java.util.ArrayList;


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
        teacher.setWorkYear(RandomUtil.randomInt(1, 10));
        teacher.setTeacherName(RandomUtil.randomString(5));
        return teacher;
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
