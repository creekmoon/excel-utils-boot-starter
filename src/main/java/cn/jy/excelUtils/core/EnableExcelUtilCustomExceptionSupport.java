package cn.jy.excelUtils.core;


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
public @interface EnableExcelUtilCustomExceptionSupport {

}
