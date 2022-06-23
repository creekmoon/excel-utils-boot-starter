package cn.jy.excelUtils.example;


import cn.jy.excelUtils.config.EnableExcelUtils;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication()
@EnableExcelUtils
public class ExampleApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExampleApplication.class, args);
    }


}
