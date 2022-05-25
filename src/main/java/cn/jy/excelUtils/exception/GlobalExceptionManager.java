package cn.jy.excelUtils.exception;

import java.util.ArrayList;
import java.util.List;

/**
 * 全局异常管理器
 */
public class GlobalExceptionManager {

    public static List<ExceptionHandler> exceptionHandlers = new ArrayList<>();

    public static String getExceptionMsg(Exception unCatchException) {
        for (int i = 0; i < exceptionHandlers.size(); i++) {
            String msg = exceptionHandlers.get(i).customExceptionMessage(unCatchException);
            if (msg != null) {
                return msg;
            }
        }
        return null;
    }

}
