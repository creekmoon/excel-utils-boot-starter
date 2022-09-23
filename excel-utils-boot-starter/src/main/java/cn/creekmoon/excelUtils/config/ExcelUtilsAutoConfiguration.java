package cn.creekmoon.excelUtils.config;

import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;
import lombok.SneakyThrows;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.ImportSelector;
import org.springframework.core.type.AnnotationMetadata;

import java.util.Map;

@Configuration
public class ExcelUtilsAutoConfiguration implements ImportSelector {


    @Override
    @SneakyThrows
    public String[] selectImports(AnnotationMetadata importingClassMetadata) {
        /*获取父注解的值*/
        Map annotationAttributes = importingClassMetadata.getAnnotationAttributes(EnableExcelUtils.class.getName(), true);
        /*获取配置项 - 自定义异常*/
        for (String customExceptionName : (String[]) annotationAttributes.get("customExceptions")) {
            GlobalExceptionManager.addExceptionHandler(new GlobalExceptionManager.DefaultExcelUtilsExceptionHandler(customExceptionName));
        }
        /*获取配置项 - 最大并行导入*/
        ExcelUtilsConfig.DEFAULT_IMPORT_MAX_PARALLEL = (Integer) annotationAttributes.get("importMaxParallel");
        /*获取配置项 - 临时文件寿命*/
        ExcelUtilsConfig.DEFAULT_TEMP_FILE_LIFE_MINUTES = (Integer) annotationAttributes.get("tempFileLifeMinutes");

        return new String[]{ExcelUtilsConfig.class.getName()};
    }
}
