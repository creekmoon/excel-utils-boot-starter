package cn.creekmoon.excel.core.R.converter;

import java.math.BigDecimal;

/**
 * @author JY
 */
public class BigDecimalConverter {


    /**
     * 浮点数值转换器
     *
     * @param str 119.12 或  1,112.99
     * @return
     */
    public static BigDecimal parse(String str) {
        if (str.contains(",")) {
            str = str.replaceAll(",", "");
        }
        return new BigDecimal(str);
    }
}
