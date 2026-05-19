package cn.creekmoon.excel.support.fixture;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDateTime;
import java.util.Date;

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
}
