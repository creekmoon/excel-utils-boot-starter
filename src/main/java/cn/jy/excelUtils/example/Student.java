package cn.jy.excelUtils.example;

import lombok.Data;

import java.time.LocalDateTime;
import java.util.Date;

@Data
public class Student {


    String userName;
    String fullName;
    String email;
    Integer age;
    Date birthday;
    LocalDateTime expTime;
}
