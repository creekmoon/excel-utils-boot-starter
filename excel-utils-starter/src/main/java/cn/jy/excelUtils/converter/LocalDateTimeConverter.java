package cn.jy.excelUtils.converter;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.date.LocalDateTimeUtil;

import java.time.LocalDateTime;
import java.util.Date;

/**
 * @author JY
 */
public class LocalDateTimeConverter {


    public static LocalDateTime parse(String str) {
        return LocalDateTimeUtil.of(DateConverter.parse(str));
    }
}
