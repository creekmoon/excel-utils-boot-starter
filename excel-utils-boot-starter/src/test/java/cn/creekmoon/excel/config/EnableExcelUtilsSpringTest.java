package cn.creekmoon.excel.config;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.support.spring.ExcelTestApplication;
import cn.creekmoon.excel.support.spring.TestBizException;
import cn.creekmoon.excel.util.exception.GlobalExceptionMsgManager;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * @EnableExcelUtils → ExcelUtilsAutoConfiguration 集成契约：
 * 注解参数（customExceptions / importMaxParallel / tempFileLifeMinutes）真正生效。
 */
@SpringBootTest(classes = ExcelTestApplication.class)
class EnableExcelUtilsSpringTest {

    @Test
    void enableExcelUtils_annotationAppliedToConfig() {
        // importMaxParallel=2 直接赋值为信号量许可数
        assertEquals(2, ExcelUtilsConfig.importParallelSemaphore.availablePermits());

        // tempFileLifeMinutes=1 与当前值取 max（AutoConfiguration 语义），至少为 1
        assertTrue(ExcelUtilsConfig.TEMP_FILE_LIFE_MINUTES >= 1);

        // customExceptions = {TestBizException.class} 已注册
        assertTrue(GlobalExceptionMsgManager.isCustomException(new TestBizException("spring-integration")),
                "TestBizException 应被注册为自定义异常");
    }
}
