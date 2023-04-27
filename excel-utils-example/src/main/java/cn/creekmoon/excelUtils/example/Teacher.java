package cn.creekmoon.excelUtils.example;

import lombok.*;


/**
 * 注意导出的对象不能重写 equal和hashCode方法
 */
@Setter
@Getter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class Teacher {


    String teacherName;
    Integer workYear;

}
