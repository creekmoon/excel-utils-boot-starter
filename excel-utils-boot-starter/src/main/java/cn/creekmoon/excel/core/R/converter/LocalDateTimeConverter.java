package cn.creekmoon.excel.core.R.converter;

import cn.creekmoon.excel.util.exception.CheckedExcelException;

import java.time.LocalDateTime;
import java.time.ZoneId;

/**
 * @author JY
 */
public class LocalDateTimeConverter {


    public static LocalDateTime parse(String str) throws CheckedExcelException {
        return LocalDateTime.ofInstant(DateConverter.parse(str).toInstant(), ZoneId.systemDefault());
    }
}
