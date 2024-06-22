package cn.creekmoon.excel.core.R.converter;

import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.hutool.core.date.LocalDateTimeUtil;

import java.time.LocalDateTime;

/**
 * @author JY
 */
public class LocalDateTimeConverter {


    public static LocalDateTime parse(String str) throws CheckedExcelException {
        return LocalDateTimeUtil.of(DateConverter.parse(str));
    }
}
