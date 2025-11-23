package cn.creekmoon.excel.util.exception;

import cn.creekmoon.excel.util.ExcelConstants;
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
public class GlobalExceptionMsgManager {

    public static final String MSG_PREFIX = "[失败!]";
    public static final String MSG_SUFFIX = "";
    public static List<CustomExceptionHandler> excelUtilsExceptionHandlers = new ArrayList<>();

    static {
        addExceptionHandler(new CustomExceptionHandler(CheckedExcelException.class));
    }


    public static boolean isCustomException(Exception exception) {
        for (CustomExceptionHandler customExceptionHandler : excelUtilsExceptionHandlers) {
            if (customExceptionHandler.isCustomException(exception)) {
                return true;
            }
        }
        return false;
    }

    public static String getExceptionMsg(Exception exception) {
        for (CustomExceptionHandler excelUtilsExceptionHandler : excelUtilsExceptionHandlers) {
            String msg = excelUtilsExceptionHandler.customExceptionMessage(exception);
            if (msg != null) {
                //可能有嵌套调用的情况 所以每次进来先把两个前缀后缀替换了 否则会出现重复的词语
                return MSG_PREFIX + msg.replace(MSG_PREFIX, "").replace(MSG_SUFFIX, "") + MSG_SUFFIX;
            }
        }
        log.error("ExcelUtils组件遇到无法处理的异常!", exception);
        return ExcelConstants.ERROR_MSG;
    }

    public static void addExceptionHandler(CustomExceptionHandler handler) {
        excelUtilsExceptionHandlers.add(handler);
        /*按优先级排序*/
        excelUtilsExceptionHandlers.sort(Comparator.comparing(CustomExceptionHandler::getOrder));
    }



}
