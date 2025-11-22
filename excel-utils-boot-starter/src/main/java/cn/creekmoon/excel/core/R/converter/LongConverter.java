package cn.creekmoon.excel.core.R.converter;

import cn.creekmoon.excel.util.exception.CheckedExcelException;
import lombok.extern.slf4j.Slf4j;

/**
 * @author JY
 */
@Slf4j
public class LongConverter {


    /**
     * 整型数值转换器
     *
     * @param str 例如 119.0   1,112.01
     * @return
     */
    public static Long parse(String str) throws CheckedExcelException {
        try {
            if (str.contains(",")) {
                str = str.replaceAll(",", "");
            }
            return str.contains(".") ? Long.valueOf(str.substring(0, str.indexOf("."))) : Long.valueOf(str);
        } catch (Exception e) {
            log.warn("[LongConverter]数据格式解析失败! 不支持的值 {}", str, e);
            throw new CheckedExcelException("不支持的数据格式!");
        }
    }
}
