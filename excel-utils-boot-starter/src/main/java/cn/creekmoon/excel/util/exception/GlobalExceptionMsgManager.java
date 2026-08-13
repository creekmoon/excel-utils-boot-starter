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
                return decorateMsg(msg);
            }
        }
        log.error("ExcelUtils组件遇到无法处理的异常!", exception);
        return ExcelConstants.ERROR_MSG;
    }

    /**
     * 为原始消息施加统一的失败前后缀
     * 可能有嵌套调用的情况, 先剥离已有前后缀再包裹, 否则会出现重复的词语
     *
     * @param rawMsg 原始消息文本, 通常取自异常的 getMessage() 或字段级失败明细
     * @return 施加前后缀后的展示消息
     */
    public static String decorateMsg(String rawMsg) {
        return MSG_PREFIX + rawMsg.replace(MSG_PREFIX, "").replace(MSG_SUFFIX, "") + MSG_SUFFIX;
    }

    public static void addExceptionHandler(CustomExceptionHandler handler) {
        excelUtilsExceptionHandlers.add(handler);
        /*按优先级排序*/
        excelUtilsExceptionHandlers.sort(Comparator.comparing(CustomExceptionHandler::getOrder));
    }



}
