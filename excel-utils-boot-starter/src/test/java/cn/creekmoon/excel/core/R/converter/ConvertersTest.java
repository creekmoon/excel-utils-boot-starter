package cn.creekmoon.excel.core.R.converter;

import cn.creekmoon.excel.util.exception.CheckedExcelException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 内置转换器契约：日期/数字/字符串的解析语义。
 *
 * <p>重点覆盖 DateConverter：Excel 序列号（1899-12-30 起算）、纯数字日期、
 * 毫秒时间戳、分隔符变体、非法输入。期望值用 java.time 独立复算，避免与实现同源。</p>
 */
class ConvertersTest {

    /**
     * Excel 序列号 → 日期时间文本的独立复算。
     * 注意：不经过 Date↔LocalDateTime 时区转换——java.time 与 java.util.TimeZone
     * 对 1900 年前的上海历史偏移不一致（+8:05:43 vs +8:00），序列号断言须用纯文本比较。
     */
    private static String serialToDateTimeText(String serial) {
        String[] split = serial.split("\\.");
        LocalDateTime result = LocalDateTime.of(1899, 12, 30, 0, 0)
                .plus(Long.parseLong(split[0]), ChronoUnit.DAYS);
        if (split.length == 2) {
            long millis = new BigDecimal("0." + split[1])
                    .multiply(BigDecimal.valueOf(24))
                    .multiply(BigDecimal.valueOf(60))
                    .multiply(BigDecimal.valueOf(60))
                    .multiply(BigDecimal.valueOf(1000))
                    .longValue();
            result = result.plus(millis, ChronoUnit.MILLIS);
        }
        return result.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
    }

    private static String fmt(Date date) {
        return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
    }

    @Nested
    @DisplayName("DateConverter")
    class DateConverterTest {

        @Test
        void excelSerialNumber_withFraction_parsesToExactDatetime() throws Exception {
            assertEquals(serialToDateTimeText("44738.94978844908"),
                    fmt(DateConverter.parse("44738.94978844908")));
        }

        @Test
        void excelSerialNumber_integerOnly_parsesToMidnight() throws Exception {
            assertEquals(serialToDateTimeText("44738"), fmt(DateConverter.parse("44738")));
        }

        @Test
        void excelSerialNumber_halfDay_addsTwelveHours() throws Exception {
            assertEquals(serialToDateTimeText("0.5"), fmt(DateConverter.parse("0.5")));
        }

        @Test
        void excelSerialNumber_dayOne_is1899_12_31() throws Exception {
            assertEquals(serialToDateTimeText("1"), fmt(DateConverter.parse("1")));
        }

        @Test
        void standardDatetimePattern_parses() throws Exception {
            assertEquals("2023-01-15 10:30:00", fmt(DateConverter.parse("2023-01-15 10:30:00")));
        }

        @Test
        void minuteOnlyPattern_parsesWithZeroSeconds() throws Exception {
            assertEquals("2023-01-15 10:30:00", fmt(DateConverter.parse("2023-01-15 10:30")));
        }

        @Test
        void dateOnlyPattern_parsesToMidnight() throws Exception {
            assertEquals("2023-01-15 00:00:00", fmt(DateConverter.parse("2023-01-15")));
        }

        @Test
        void slashSeparator_equivalentToDash() throws Exception {
            assertEquals("2023-01-15 00:00:00", fmt(DateConverter.parse("2023/01/15")));
        }

        @Test
        void dotSeparator_equivalentToDash() throws Exception {
            assertEquals("2023-01-15 00:00:00", fmt(DateConverter.parse("2023.01.15")));
        }

        @Test
        void singleDigitMonthAndDay_parses() throws Exception {
            assertEquals("2023-02-03 00:00:00", fmt(DateConverter.parse("2023-2-3")));
        }

        @Test
        void fractionSeconds_truncatedToMillis() throws Exception {
            assertEquals("2023-01-15 10:30:00.123", new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS")
                    .format(DateConverter.parse("2023-01-15 10:30:00.123456")));
        }

        @Test
        void thirteenDigitValue_throwsAsSerialOutOfRange() {
            // 现状契约：任何纯数字一律按 Excel 序列号天数处理（hutool 时代行为），
            // 13 位时间戳换算天数超出 LocalDateTime 范围 → 解析失败
            assertThrows(CheckedExcelException.class, () -> DateConverter.parse("1673764200000"));
        }

        @Test
        void pureNumericString_treatedAsSerialNotDate() {
            // "20230115" 不会被当作 yyyyMMdd 日期，而是按序列号 20230115 天换算（约 57288 年）
            try {
                String result = fmt(DateConverter.parse("20230115"));
                assertTrue(!result.startsWith("2023"), "不应按 2023-01-15 解析: " + result);
            } catch (CheckedExcelException expected) {
                // 序列号换算结果超出 SimpleDateFormat 可表示范围时抛异常，同样是"非日期"语义
            }
        }

        @Test
        void surroundingWhitespace_isTrimmed() throws Exception {
            assertEquals("2023-01-15 00:00:00", fmt(DateConverter.parse("  2023-01-15  ")));
        }

        @Test
        void isoLocalDateTime_withT_parses() throws Exception {
            // LocalDateTime.toString 输出（EasyExcel 读回 LocalDateTime 单元格后 String.valueOf 的形态）
            assertEquals("2024-06-01 12:00:00", fmt(DateConverter.parse("2024-06-01T12:00")));
            assertEquals("2024-06-01 12:00:30", fmt(DateConverter.parse("2024-06-01T12:00:30")));
        }

        @Test
        void isoLocalDateTime_withFractionalSeconds_truncatedToMillis() throws Exception {
            assertEquals("2024-06-01 12:00:30.123", new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS")
                    .format(DateConverter.parse("2024-06-01T12:00:30.123456789")));
        }

        @Test
        void illegalValue_throwsCheckedExcelException() {
            assertThrows(CheckedExcelException.class, () -> DateConverter.parse("abc"));
        }

        @Test
        void emptyValue_throwsCheckedExcelException() {
            assertThrows(CheckedExcelException.class, () -> DateConverter.parse(""));
        }

        @Test
        void futureDateAndPre1900Date_parse() throws Exception {
            assertEquals("2099-12-31 23:59:59", fmt(DateConverter.parse("2099-12-31 23:59:59")));
            assertEquals("1899-01-01 00:00:00", fmt(DateConverter.parse("1899-01-01")));
        }
    }

    @Nested
    @DisplayName("LocalDateTimeConverter")
    class LocalDateTimeConverterTest {

        @Test
        void standardDatetime_parsesToLocalDateTime() throws Exception {
            LocalDateTime expected = LocalDateTime.of(2023, 1, 15, 10, 30, 0);
            assertEquals(expected, LocalDateTimeConverter.parse("2023-01-15 10:30:00"));
        }

        @Test
        void dateOnly_parsesToLocalDateTimeMidnight() throws Exception {
            assertEquals(LocalDateTime.of(2023, 1, 15, 0, 0, 0),
                    LocalDateTimeConverter.parse("2023-01-15"));
        }

        @Test
        void excelSerialNumber_parsesThroughDateConverter() throws Exception {
            // 按序列号语义直接复算 LocalDateTime（0.5 天 = 12 小时）
            LocalDateTime expected = LocalDateTime.of(1899, 12, 30, 0, 0)
                    .plusDays(44738).plus(12, ChronoUnit.HOURS);
            assertEquals(expected, LocalDateTimeConverter.parse("44738.5"));
        }
    }

    @Nested
    @DisplayName("IntegerConverter")
    class IntegerConverterTest {

        @Test
        void plainInteger_parses() throws Exception {
            assertEquals(42, IntegerConverter.parse("42"));
        }

        @Test
        void negativeInteger_parses() throws Exception {
            assertEquals(-5, IntegerConverter.parse("-5"));
        }

        @Test
        void decimalValue_truncatesFraction() throws Exception {
            assertEquals(119, IntegerConverter.parse("119.0"));
            assertEquals(119, IntegerConverter.parse("119.99"));
        }

        @Test
        void thousandSeparator_removed() throws Exception {
            assertEquals(1112, IntegerConverter.parse("1,112.01"));
            assertEquals(1112, IntegerConverter.parse("1,112"));
        }

        @Test
        void illegalValue_throwsCheckedExcelException() {
            assertThrows(CheckedExcelException.class, () -> IntegerConverter.parse("abc"));
            assertThrows(CheckedExcelException.class, () -> IntegerConverter.parse("12a"));
            assertThrows(CheckedExcelException.class, () -> IntegerConverter.parse(""));
        }
    }

    @Nested
    @DisplayName("LongConverter")
    class LongConverterTest {

        @Test
        void maxLong_parses() throws Exception {
            assertEquals(Long.MAX_VALUE, LongConverter.parse("9223372036854775807"));
        }

        @Test
        void decimalValue_truncatesFraction() throws Exception {
            assertEquals(9223372036854775807L, LongConverter.parse("9223372036854775807.9"));
        }

        @Test
        void thousandSeparator_removed() throws Exception {
            assertEquals(1112L, LongConverter.parse("1,112.99"));
        }

        @Test
        void illegalValue_throwsCheckedExcelException() {
            assertThrows(CheckedExcelException.class, () -> LongConverter.parse("abc"));
        }
    }

    @Nested
    @DisplayName("BigDecimalConverter")
    class BigDecimalConverterTest {

        @Test
        void decimalValue_preservesPrecision() {
            assertEquals(new BigDecimal("119.12"), BigDecimalConverter.parse("119.12"));
        }

        @Test
        void thousandSeparator_removed() {
            assertEquals(new BigDecimal("1112.99"), BigDecimalConverter.parse("1,112.99"));
        }

        @Test
        void smallFraction_preserved() {
            assertEquals(new BigDecimal("0.0000001"), BigDecimalConverter.parse("0.0000001"));
        }

        @Test
        void scientificNotation_supported() {
            assertEquals(new BigDecimal("1E+5"), BigDecimalConverter.parse("1E5"));
        }

        @Test
        void illegalValue_throwsNumberFormatException() {
            // 现状契约：BigDecimalConverter 不包装异常，非法值直接抛 NumberFormatException
            assertThrows(NumberFormatException.class, () -> BigDecimalConverter.parse("abc"));
        }
    }

    @Nested
    @DisplayName("StringConverter")
    class StringConverterTest {

        @Test
        void nullValue_returnsEmptyString() {
            assertEquals("", StringConverter.parse(null));
        }

        @Test
        void stringValue_isTrimmed() {
            assertEquals("abc", StringConverter.parse("  abc  "));
        }

        @Test
        void dateValue_formattedToStandardPattern() {
            Date date = Date.from(LocalDateTime.of(2024, 6, 1, 12, 0, 0)
                    .atZone(ZoneId.systemDefault()).toInstant());
            assertEquals("2024-06-01 12:00:00", StringConverter.parse(date));
        }

        @Test
        void numericValue_preservedAsStringValueOf() {
            assertEquals("123.0", StringConverter.parse(123.0));
            assertEquals("42", StringConverter.parse(42));
        }

        @Test
        void unknownObject_usesToString() {
            assertEquals("x", StringConverter.parse(new Object() {
                @Override
                public String toString() {
                    return "x";
                }
            }));
        }
    }
}
