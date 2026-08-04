package cn.creekmoon.excel.core.R.converter;


import cn.creekmoon.excel.util.exception.CheckedExcelException;
import lombok.extern.slf4j.Slf4j;

import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.Date;

/**
 * @author JY
 */
@Slf4j
public class DateConverter {

    public static final DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final BigDecimal BIG_DECIMAL_24 = new BigDecimal(24);
    private static final BigDecimal BIG_DECIMAL_60 = new BigDecimal(60);
    private static final BigDecimal BIG_DECIMAL_1000 = new BigDecimal(1000);
    /* 重要 PS:EXCEL的起始时间是1899-12-30开始算的 而不是1900-1-1 不要问我为什么,我也不知道 */
    private static final LocalDateTime EXCEL_START_TIME = LocalDateTime.of(1899, 12, 30, 0, 0);

    public static Date parse(String dateStr) throws CheckedExcelException {

        try {
            //针对EXCEL传过来的时间进行处理 例如44738.94978844908  代表44738.94978844908天
            if (isNumber(dateStr)) {
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


            return parseDate(dateStr);
        } catch (Exception e) {
            log.warn("[DateConverter]数据格式解析失败! 不支持的值 {}", dateStr, e);
            throw new CheckedExcelException("不支持的数据格式!");
        }
    }

    /**
     * 判断字符串是否为数字(兼容小数、科学计数法),行为等价于 Hutool NumberUtil.isNumber
     */
    private static boolean isNumber(String str) {
        try {
            new BigDecimal(str);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * 使用 JDK SimpleDateFormat 解析日期字符串,行为兼容 Hutool DateUtil.parse 的常用格式:
     * <ul>
     *   <li>标准格式:yyyy-MM-dd HH:mm:ss / HH:mm / 日期部分,分隔符 - / . 等价</li>
     *   <li>ISO 格式(LocalDateTime.toString 输出,含 T):T 视为空格,毫秒保留</li>
     *   <li>纯数字:13位按毫秒时间戳,8/14/17位按纯数字日期时间</li>
     * </ul>
     */
    private static Date parseDate(String dateStr) throws ParseException {
        String str = dateStr.trim();
        /* 纯数字形式 */
        if (str.matches("\\d+")) {
            if (str.length() == 13) {
                return new Date(Long.parseLong(str));
            }
            return parseByPatterns(str, "yyyyMMddHHmmssSSS", "yyyyMMddHHmmss", "yyyyMMdd");
        }
        String normalized;
        int msIndex = str.lastIndexOf('.');
        if (msIndex > 0 && msIndex > str.lastIndexOf(':') && str.length() - msIndex > 4) {
            /* 时间(冒号)后的 "." 是毫秒分隔:截断到 3 位并保留,否则会被 SimpleDateFormat 静默丢弃 */
            normalized = str.substring(0, msIndex).replace('T', ' ').replace('/', '-').replace('.', '-')
                    + str.substring(msIndex, msIndex + 4);
        } else {
            /* ISO 格式(LocalDateTime.toString 输出,如 2024-06-01T12:00:30.123456789):
             * 含 T 时毫秒小数点需保留,不再做分隔符替换 */
            normalized = str.contains("T") ? str.replace('T', ' ')
                    : str.replace('/', '-').replace('.', '-');
        }
        return parseByPatterns(normalized,
                "yyyy-MM-dd HH:mm:ss.SSS", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM-dd");
    }

    private static Date parseByPatterns(String str, String... patterns) throws ParseException {
        for (String pattern : patterns) {
            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
            sdf.setLenient(false);
            try {
                return sdf.parse(str);
            } catch (ParseException ignored) {
                /* 尝试下一个格式 */
            }
        }
        throw new ParseException("无法解析的日期字符串: " + str, 0);
    }
}
