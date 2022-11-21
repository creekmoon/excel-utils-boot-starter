package cn.creekmoon.excelUtils.exception;

import cn.creekmoon.excelUtils.core.ExcelConstants;
import lombok.SneakyThrows;
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

    public static final String MSG_PREFIX = "导入失败!";
    public static final String MSG_SUFFIX = "";
    public static List<ExcelUtilsExceptionHandler> excelUtilsExceptionHandlers = new ArrayList<>();

    static {
        addExceptionHandler(new DefaultExcelUtilsExceptionHandler(CheckedExcelException.class));
    }

    public static String getExceptionMsg(Exception unCatchException) {
        for (ExcelUtilsExceptionHandler excelUtilsExceptionHandler : excelUtilsExceptionHandlers) {
            String msg = excelUtilsExceptionHandler.customExceptionMessage(unCatchException);
            if (msg != null) {
                //可能有嵌套调用的情况 所以每次进来先把两个前缀后缀替换了 否则会出现重复的词语
                return MSG_PREFIX + msg.replace(MSG_PREFIX, "").replace(MSG_SUFFIX, "") + MSG_SUFFIX;
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

    /**
     * 默认异常处理器
     */
    public static class DefaultExcelUtilsExceptionHandler implements ExcelUtilsExceptionHandler {

        /*自定义异常*/
        private Class customExceptionClass = null;

        public DefaultExcelUtilsExceptionHandler(Class customExceptionClass) {
            this.customExceptionClass = customExceptionClass;
        }

        @SneakyThrows
        public DefaultExcelUtilsExceptionHandler(String customExceptionClassName) {
            this.customExceptionClass = Class.forName(customExceptionClassName);
        }

        @Override
        public String customExceptionMessage(Exception unCatchException) {
            if (customExceptionClass.isInstance(unCatchException)) {
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
