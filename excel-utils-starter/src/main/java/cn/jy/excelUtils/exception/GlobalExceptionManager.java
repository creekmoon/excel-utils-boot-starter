package cn.jy.excelUtils.exception;

import cn.jy.excelUtils.core.ExcelConstants;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.Ordered;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * 全局异常管理器
 */
@Slf4j
public class GlobalExceptionManager {

    public static List<ExcelUtilsExceptionHandler> excelUtilsExceptionHandlers = new ArrayList<>();

    static {
        addExceptionHandler(new DefaultExcelUtilsExceptionHandler());
    }
    public static String getExceptionMsg(Exception unCatchException) {
        for (ExcelUtilsExceptionHandler excelUtilsExceptionHandler : excelUtilsExceptionHandlers) {
            String msg = excelUtilsExceptionHandler.customExceptionMessage(unCatchException);
            if (msg != null) {
                return msg;
            }
        }
        log.error("ExcelUtils组件遇到无法处理的异常!", unCatchException);
        return ExcelConstants.ERROR_MSG;
    }

    public static void addExceptionHandler(ExcelUtilsExceptionHandler handler) {
        excelUtilsExceptionHandlers.add(handler);
        /*按优先级排序*/
        excelUtilsExceptionHandlers.sort(Comparator.comparing(ExcelUtilsExceptionHandler::getOrder));
    }

    public static class DefaultExcelUtilsExceptionHandler implements ExcelUtilsExceptionHandler {
        @Override
        public String customExceptionMessage(Exception unCatchException) {
            if (unCatchException instanceof CheckedExcelException) {
                return unCatchException.getMessage();
            }
            //返回null说明不进行处理 将委托给其他处理器进行处理
            return null;
        }

        @Override
        public int getOrder() {
            return Ordered.LOWEST_PRECEDENCE;
        }
    }
}
