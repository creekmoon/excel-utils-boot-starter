package cn.creekmoon.excelUtils.example;


import cn.creekmoon.excelUtils.config.EnableExcelUtils;
import cn.creekmoon.excelUtils.example.config.exception.MyNewException;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication()
@EnableExcelUtils(customExceptions = {MyNewException.class})
public class ExampleApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExampleApplication.class, args);
    }


}
