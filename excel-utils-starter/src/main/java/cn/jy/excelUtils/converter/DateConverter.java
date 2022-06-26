package cn.jy.excelUtils.converter;

import cn.hutool.core.date.DateTime;
import cn.hutool.core.date.DateUtil;

import java.util.Date;

/**
 * @author JY
 */
public class DateConverter {


    public static Date parse(String str) {
        return DateUtil.parse(str);
    }
}
