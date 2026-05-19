package cn.creekmoon.excel.support.spring;

import org.springframework.boot.SpringBootConfiguration;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;

/** 测试用 Spring 配置壳（当前用例多直接改 ExcelUtilsConfig，保留供 @SpringBootTest 扩展）。 */
@SpringBootConfiguration
@EnableAutoConfiguration
@EnableExcelUtils(customExceptions = {TestBizException.class}, importMaxParallel = 2, tempFileLifeMinutes = 1)
public class ExcelTestApplication {
}
