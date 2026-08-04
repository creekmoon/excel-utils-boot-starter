package cn.creekmoon.excel.util;

/**
 * Excel常量
 */
public class ExcelConstants {
    public static final String ERROR_MSG = "导入异常!请联系管理员!";
    public static final String RESULT_TITLE = "导入结果";
    public static final String IMPORT_SUCCESS_MSG = "导入成功!";
    public static final String CONVERT_SUCCESS_MSG = "验证通过!";
    /* 注意:占位符使用 %s(String.format 风格),替代 Hutool StrFormatter 的 {} 语法 */
    public static final String CONVERT_FAIL_MSG = "字段[%s]解析失败！";

    public static final String FIELD_LACK_MSG = "字段[%s]为必填项! ";

    /**
     * 日期单元格统一读取格式。
     * EasyExcel 在无实体类(Map 模式)下默认把日期单元格转成分钟精度字符串，
     * 需通过 readBuilder.dateFormat() 显式指定，避免秒级信息丢失。
     */
    public static final String DATE_FORMAT_PATTERN = "yyyy-MM-dd HH:mm:ss";

}
