package cn.creekmoon.excelUtils.config;


import org.springframework.context.annotation.Import;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 对于导入导出组件,启用自定义异常处理器的支持
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Import({ExcelUtilsAutoConfiguration.class})
public @interface EnableExcelUtils {

    Class[] customExceptions() default {};

    /**
     * 能并行执行多少个导入任务 防止内存溢出 初始化后会赋值给ExcelUtilsConfig
     */
    int importMaxParallel() default 4;

    /**
     * 临时文件的保留寿命 单位分钟  初始化后会赋值给ExcelUtilsConfig
     */
    int tempFileLifeMinutes() default 5;
}
