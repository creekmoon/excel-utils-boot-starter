package cn.creekmoon.excelUtils.config;

import cn.creekmoon.excelUtils.core.ExcelImport;
import cn.creekmoon.excelUtils.exception.ExcelUtilsExceptionHandler;
import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;
import cn.creekmoon.excelUtils.threadPool.CleanTempFilesExecutor;
import lombok.Data;
import org.springframework.beans.BeansException;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;
import java.util.Collection;
import java.util.Comparator;
import java.util.concurrent.Semaphore;

/**
 * ES连接常量
 */
@Component //定义配置类
@Data //提供get set方法
@ConfigurationProperties(prefix = "excel-utils") //yml配置中的路径
public class ExcelUtilsConfig implements ApplicationContextAware {

    public static int DEFAULT_IMPORT_MAX_PARALLEL = 4;
    public static int DEFAULT_TEMP_FILE_LIFE_MINUTES = 4;

    /**
     * 能并行执行多少个导入任务 防止内存溢出
     */
    private int importMaxParallel = 0;

    /**
     * 临时文件的保留寿命 单位分钟
     */
    private int tempFileLifeMinutes = 0;

    ApplicationContext applicationContext;


    @Override
    public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
        this.applicationContext = applicationContext;
    }

    @PostConstruct
    public void init() {
        /*配置默认值*/
        if (importMaxParallel <= 0) {
            importMaxParallel = DEFAULT_IMPORT_MAX_PARALLEL;
        }
        if (tempFileLifeMinutes <= 0) {
            tempFileLifeMinutes = DEFAULT_TEMP_FILE_LIFE_MINUTES;
        }
        /*如果用户实现了ExcelUtilsExceptionHandler接口, 装配用户自己实现的异常处理器*/
        Collection<ExcelUtilsExceptionHandler> values = applicationContext.getBeansOfType(ExcelUtilsExceptionHandler.class).values();
        GlobalExceptionManager.excelUtilsExceptionHandlers.addAll(values);
        GlobalExceptionManager.excelUtilsExceptionHandlers.sort(Comparator.comparing(ExcelUtilsExceptionHandler::getOrder));
        /*装配用户定义的参数*/
        CleanTempFilesExecutor.TEMP_FILE_LIFE_MINUTES = importMaxParallel;
        CleanTempFilesExecutor.init();
        ExcelImport.importSemaphore = new Semaphore(tempFileLifeMinutes);
    }
}
