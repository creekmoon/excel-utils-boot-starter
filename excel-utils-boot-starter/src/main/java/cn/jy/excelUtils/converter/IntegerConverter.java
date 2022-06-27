package cn.jy.excelUtils.converter;

/**
 * @author JY
 */
public class IntegerConverter {


    public static Integer parse(String str) {
        return str.contains(".") ? Integer.valueOf(str.substring(0, str.indexOf("."))) : Integer.valueOf(str);
    }
}
