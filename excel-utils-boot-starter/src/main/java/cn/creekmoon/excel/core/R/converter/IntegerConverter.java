package cn.creekmoon.excel.core.R.converter;

import cn.creekmoon.excel.util.exception.CheckedExcelException;
import lombok.extern.slf4j.Slf4j;

/**
 * @author JY
 */
@Slf4j
public class IntegerConverter {


    /**
     * 整型数值转换器
     *
     * @param str 例如 119.0   1,112.01
     * @return
     */
    public static Integer parse(String str) throws CheckedExcelException {
        try {
            str = str.replaceAll(",", "");
            return str.contains(".") ? Integer.valueOf(str.substring(0, str.indexOf("."))) : Integer.valueOf(str);
        } catch (Exception e) {
            log.warn("[IntegerConverter]数据格式解析失败! 不支持的值 {}", str, e);
            throw new CheckedExcelException("不支持的数据格式!");
        }
    }
}
