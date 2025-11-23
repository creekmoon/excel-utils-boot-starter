package cn.creekmoon.excel.support.spring;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.util.exception.CustomExceptionHandler;
import cn.creekmoon.excel.util.exception.GlobalExceptionMsgManager;
import lombok.SneakyThrows;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.ImportSelector;
import org.springframework.core.type.AnnotationMetadata;

import java.util.Map;
import java.util.concurrent.Semaphore;

@Configuration
public class ExcelUtilsAutoConfiguration implements ImportSelector {


    @Override
    @SneakyThrows
    public String[] selectImports(AnnotationMetadata importingClassMetadata) {
        /*获取父注解的值*/
        Map annotationAttributes = importingClassMetadata.getAnnotationAttributes(EnableExcelUtils.class.getName(), true);
        /*获取配置项 - 自定义异常*/
        for (String customExceptionName : (String[]) annotationAttributes.get("customExceptions")) {
            ExcelUtilsConfig.addExcelUtilsExceptionHandler(new CustomExceptionHandler(customExceptionName));
        }
        /*获取配置项 - 最大并行导入*/
        ExcelUtilsConfig.importParallelSemaphore = new Semaphore(Math.max((Integer) annotationAttributes.get("importMaxParallel"), 1));
        /*获取配置项 - 临时文件寿命*/
        ExcelUtilsConfig.TEMP_FILE_LIFE_MINUTES = Math.max((Integer) annotationAttributes.get("tempFileLifeMinutes"), ExcelUtilsConfig.TEMP_FILE_LIFE_MINUTES);

        return new String[]{ExcelUtilsConfig.class.getName()};
    }
}
