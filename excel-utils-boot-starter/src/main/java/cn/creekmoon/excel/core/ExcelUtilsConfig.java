package cn.creekmoon.excel.core;

import cn.creekmoon.excel.util.exception.CustomExceptionHandler;
import cn.creekmoon.excel.util.exception.GlobalExceptionMsgManager;
import lombok.Data;

import java.util.Arrays;
import java.util.Comparator;
import java.util.concurrent.Semaphore;

/**
 * 导出导入工具类配置常量
 */
@Data //提供get set方法
public class ExcelUtilsConfig {

    /**
     * 控制导入并发数量
     */
    public static Semaphore importParallelSemaphore = new Semaphore(4);

    /**
     * 临时文件的保留寿命 单位分钟
     */
    public static int TEMP_FILE_LIFE_MINUTES = 5;


    public static void addExcelUtilsExceptionHandler(CustomExceptionHandler... handlers) {
        if (handlers == null || handlers.length == 0) {
            return;
        }
        GlobalExceptionMsgManager.excelUtilsExceptionHandlers.addAll(Arrays.asList(handlers));
        GlobalExceptionMsgManager.excelUtilsExceptionHandlers.sort(Comparator.comparing(CustomExceptionHandler::getOrder));
    }

}
