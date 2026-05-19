package cn.creekmoon.excel.support.fixture;

import java.util.ArrayList;

public final class TeacherFixtures {

    private TeacherFixtures() {
    }

    public static Teacher createDeterministic(int index) {
        Teacher teacher = new Teacher();
        teacher.setTeacherName("teacher" + index);
        teacher.setWorkYear(1 + (index % 10));
        return teacher;
    }

    public static ArrayList<Teacher> createList(int size) {
        ArrayList<Teacher> result = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            result.add(createDeterministic(i));
        }
        return result;
    }
}
