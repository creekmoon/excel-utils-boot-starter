package cn.creekmoon.excelUtils.core;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.function.Supplier;

/**
 * Excel常量
 */
public class ExcelConstants {
    public static final String ERROR_MSG = "导入异常!请联系管理员!";
    public static final String RESULT_TITLE = "导入结果";
    public static final String IMPORT_SUCCESS_MSG = "导入成功!";
    public static final String CONVERT_SUCCESS_MSG = "验证通过!";
    public static final String CONVERT_FAIL_MSG = "字段[{}]解析失败！";

    public static final String FIELD_LACK_MSG = "字段[{}]为必填项! 请检查格式！";

    public static final String TITLE_CHECK_ERROR = "导入的模板有误,请检查您的文件!";

    public static final Supplier<String> excelNameGenerator = () -> "export_result_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy_MM_dd_HH_mm"));
}
