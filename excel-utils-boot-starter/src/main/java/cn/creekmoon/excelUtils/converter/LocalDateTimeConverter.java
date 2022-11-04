package cn.creekmoon.excelUtils.converter;

import cn.creekmoon.excelUtils.hutool589.core.date.LocalDateTimeUtil;

import java.time.LocalDateTime;

/**
 * @author JY
 */
public class LocalDateTimeConverter {


    public static LocalDateTime parse(String str) {
        return LocalDateTimeUtil.of(DateConverter.parse(str));
    }
}
