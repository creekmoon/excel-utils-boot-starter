package cn.creekmoon.excelUtils.converter;

import cn.creekmoon.excelUtils.hutool589.core.date.DatePattern;
import cn.creekmoon.excelUtils.hutool589.core.date.DateUtil;
import cn.creekmoon.excelUtils.hutool589.core.util.NumberUtil;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.Date;

/**
 * @author JY
 */
public class DateConverter {

    public static final DateTimeFormatter fmt = DateTimeFormatter.ofPattern(DatePattern.NORM_DATETIME_PATTERN);
    private static final BigDecimal BIG_DECIMAL_24 = new BigDecimal(24);
    private static final BigDecimal BIG_DECIMAL_60 = new BigDecimal(60);
    private static final BigDecimal BIG_DECIMAL_1000 = new BigDecimal(1000);
    /* 重要 PS:EXCEL的起始时间是1899-12-30开始算的 而不是1900-1-1 不要问我为什么,我也不知道 */
    private static final LocalDateTime EXCEL_START_TIME = LocalDateTime.of(1899, 12, 30, 0, 0);

    public static Date parse(String dateStr) {


        //针对EXCEL传过来的时间进行处理 例如44738.94978844908  代表44738.94978844908天
        if (NumberUtil.isNumber(dateStr)) {
            LocalDateTime start = EXCEL_START_TIME;
            /*小数部分按毫秒级算*/
            if (dateStr.contains(".")) {
                String[] split = dateStr.split("\\.");
                dateStr = split[0];
                BigDecimal multiply = new BigDecimal("0." + split[1])
                        .multiply(BIG_DECIMAL_24)
                        .multiply(BIG_DECIMAL_60)
                        .multiply(BIG_DECIMAL_60)
                        .multiply(BIG_DECIMAL_1000);
                start = start.plus(multiply.longValue(), ChronoUnit.MILLIS);
            }
            /*整数部分按天算*/
            start = start.plus(new BigDecimal(dateStr).longValue(), ChronoUnit.DAYS);
            dateStr = start.format(fmt);
        }


        return DateUtil.parse(dateStr);
    }
}
