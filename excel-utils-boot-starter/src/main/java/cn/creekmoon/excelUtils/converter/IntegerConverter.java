package cn.creekmoon.excelUtils.converter;

/**
 * @author JY
 */
public class IntegerConverter {


    /**
     * 整型数值转换器
     *
     * @param str 例如 119.0   1,112.01
     * @return
     */
    public static Integer parse(String str) {
        str = str.replaceAll(",", "");
        return str.contains(".") ? Integer.valueOf(str.substring(0, str.indexOf("."))) : Integer.valueOf(str);
    }
}
