package cn.creekmoon.excel.support.fixture;

import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;

public final class StudentFixtures {

    private static final LocalDateTime BASE_TIME = LocalDateTime.of(2024, 6, 1, 12, 0, 0);
    private static final Date BASE_DATE = Date.from(BASE_TIME.atZone(ZoneId.systemDefault()).toInstant());

    private StudentFixtures() {
    }

    public static Student createDeterministic(int index) {
        Student student = new Student();
        student.setUserName("user" + index);
        student.setFullName("full" + index);
        student.setEmail("user" + index + "@qq.com");
        student.setAge(20 + (index % 80));
        student.setBirthday(BASE_DATE);
        student.setExpTime(BASE_TIME.plusDays(index));
        return student;
    }

    public static ArrayList<Student> createList(int size) {
        ArrayList<Student> result = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            result.add(createDeterministic(i));
        }
        return result;
    }
}
