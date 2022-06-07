package cn.jy.excelUtils.exception;

import cn.jy.excelUtils.core.ExcelConstants;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * 全局异常管理器
 */
@Slf4j
public class GlobalExceptionManager {

    public static List<ExceptionHandler> exceptionHandlers = new ArrayList<>();
    static {
        addExceptionHandler(new DefaultExceptionHandler());
    }
    public static String getExceptionMsg(Exception unCatchException) {
        for (ExceptionHandler exceptionHandler : exceptionHandlers) {
            String msg = exceptionHandler.customExceptionMessage(unCatchException);
            if (msg != null) {
                return msg;
            }
        }
        log.error("ExcelUtils组件遇到无法处理的异常!", unCatchException);
        return ExcelConstants.ERROR_MSG;
    }

    public static void addExceptionHandler(ExceptionHandler handler) {
        exceptionHandlers.add(handler);
        /*按等级从大到小排序*/
        exceptionHandlers.sort(Comparator.comparing(ExceptionHandler::getLevel).reversed());
    }
}
