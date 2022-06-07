package cn.jy.excelUtils.core;

import cn.jy.excelUtils.exception.ExcelUtilsExceptionHandler;
import cn.jy.excelUtils.exception.GlobalExceptionManager;
import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.context.annotation.Configuration;

import javax.annotation.PostConstruct;
import java.util.Collection;
import java.util.Comparator;

@Configuration
public class ExcelUtilsAutoConfigure implements ApplicationContextAware {

    ApplicationContext applicationContext;

    @PostConstruct
    public void init(){
        /*装配异常处理器*/
        Collection<ExcelUtilsExceptionHandler> values = applicationContext.getBeansOfType(ExcelUtilsExceptionHandler.class).values();
        GlobalExceptionManager.excelUtilsExceptionHandlers.addAll(values);
        GlobalExceptionManager.excelUtilsExceptionHandlers.sort(Comparator.comparing(ExcelUtilsExceptionHandler::getOrder));

    }
    @Override
    public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
        this.applicationContext = applicationContext;
    }
}
