package cn.jy.excelUtils.exception;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * 全局异常管理器
 */
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
        return null;
    }

    public static void addExceptionHandler(ExceptionHandler handler) {
        exceptionHandlers.add(handler);
        /*按等级从大到小排序*/
        exceptionHandlers.sort(Comparator.comparing(ExceptionHandler::getLevel).reversed());
    }
}
