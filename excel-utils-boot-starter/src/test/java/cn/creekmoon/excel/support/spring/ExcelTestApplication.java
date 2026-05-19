package cn.creekmoon.excel.support.spring;

import org.springframework.boot.SpringBootConfiguration;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;

@SpringBootConfiguration
@EnableAutoConfiguration
@EnableExcelUtils(customExceptions = {TestBizException.class}, importMaxParallel = 2, tempFileLifeMinutes = 1)
public class ExcelTestApplication {
}
