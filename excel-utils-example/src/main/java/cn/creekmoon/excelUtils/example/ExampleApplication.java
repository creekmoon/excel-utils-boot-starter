package cn.creekmoon.excelUtils.example;


import cn.creekmoon.excelUtils.config.EnableExcelUtils;
import cn.creekmoon.excelUtils.example.config.exception.MyNewException;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication()
//Excel导入的整个过程中,如果抛出MyNewException异常,将异常信息作为导入结果
@EnableExcelUtils(customExceptions = {MyNewException.class}, importMaxParallel = 4, tempFileLifeMinutes = 5)
public class ExampleApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExampleApplication.class, args);
    }


}
