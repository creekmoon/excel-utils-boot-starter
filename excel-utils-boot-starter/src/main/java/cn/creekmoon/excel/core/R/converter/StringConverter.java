package cn.creekmoon.excel.core.R.converter;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.StrUtil;

import java.util.Date;

/**
 * 元数据转换器
 * Excel初次解析时，就会执行这个解析器。将所有数据都转成String
 */
public class StringConverter {

    /**
     * 全局默认数据转换器 将所有的数值都转换成String
     *
     * @param value
     * @return
     */
    public static String parse(Object value) {
        if (value == null) {
            return "";
        }
        if (value instanceof String) {
            return StrUtil.trim((String) value);
        }
        if (value instanceof Date) {
            return DateUtil.format((Date) value, "yyyy-MM-dd HH:mm:ss");
        }
        return String.valueOf(value);
    }
}
