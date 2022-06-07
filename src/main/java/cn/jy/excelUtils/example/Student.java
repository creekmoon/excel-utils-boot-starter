package cn.jy.excelUtils.example;

import cn.hutool.core.text.StrFormatter;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import java.time.LocalDateTime;
import java.util.Date;


/**
 * 注意导出的对象不能重写 equal和hashCode方法
 */
@Setter
@Getter
public class Student {


    String userName;
    String fullName;
    String email;
    Integer age;
    Date birthday;
    LocalDateTime expTime;
}
