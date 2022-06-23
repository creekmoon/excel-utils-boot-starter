package cn.jy.excelUtils.config;

import cn.jy.excelUtils.core.ExcelImport;
import cn.jy.excelUtils.exception.ExcelUtilsExceptionHandler;
import cn.jy.excelUtils.exception.GlobalExceptionManager;
import cn.jy.excelUtils.threadPool.AsyncReadStateCallbackExecutor;
import org.springframework.beans.BeansException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.context.annotation.Configuration;

import javax.annotation.PostConstruct;
import java.util.Collection;
import java.util.Comparator;
import java.util.concurrent.Semaphore;

@Configuration
public class ExcelUtilsAutoConfiguration implements ApplicationContextAware {

    ApplicationContext applicationContext;

    @Autowired
    ExcelUtilsConfig excelUtilsConfig;

    @PostConstruct
    public void init() {
        /*装配异常处理器*/
        Collection<ExcelUtilsExceptionHandler> values = applicationContext.getBeansOfType(ExcelUtilsExceptionHandler.class).values();
        GlobalExceptionManager.excelUtilsExceptionHandlers.addAll(values);
        GlobalExceptionManager.excelUtilsExceptionHandlers.sort(Comparator.comparing(ExcelUtilsExceptionHandler::getOrder));
        /*初始化线程池配置*/
        AsyncReadStateCallbackExecutor.init(1, 1500);
        /*初始化最大的导入执行数量*/
        ExcelImport.semaphore = new Semaphore(excelUtilsConfig.getImportMaxParallel());
    }
    @Override
    public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
        this.applicationContext = applicationContext;
    }
}
