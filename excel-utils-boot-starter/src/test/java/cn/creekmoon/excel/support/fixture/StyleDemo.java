package cn.creekmoon.excel.support.fixture;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 * 样式展示专用数据类（复刻 excel-utils-example 的 StyleDemo）。
 */
@Setter
@Getter
@Builder
@ToString
@NoArgsConstructor
@AllArgsConstructor
public class StyleDemo {

    private String name;
    private Integer age;
    private BigDecimal amount;
    private Double rate;
    private LocalDate date;
    private LocalDateTime datetime;
    private String status;
    private String category;
    private Integer score;

    public static StyleDemo create(String name, Integer age, BigDecimal amount, Double rate,
                                    LocalDate date, LocalDateTime datetime, String status,
                                    String category, Integer score) {
        return StyleDemo.builder()
                .name(name)
                .age(age)
                .amount(amount)
                .rate(rate)
                .date(date)
                .datetime(datetime)
                .status(status)
                .category(category)
                .score(score)
                .build();
    }

    /** 确定性样式演示数据：4 行覆盖 正常/警告/异常 状态与 高/中/低 分数。 */
    public static java.util.List<StyleDemo> createShowcaseData() {
        return java.util.Arrays.asList(
                create("张三", 25, new BigDecimal("12345.67"), 0.15,
                        LocalDate.of(2024, 1, 1), LocalDateTime.of(2024, 1, 1, 10, 30), "正常", "A类", 85),
                create("李四", 45, new BigDecimal("99999.99"), 0.85,
                        LocalDate.of(2024, 1, 2), LocalDateTime.of(2024, 1, 2, 11, 30), "警告", "B类", 72),
                create("王五", 65, new BigDecimal("500.00"), 0.05,
                        LocalDate.of(2024, 1, 3), LocalDateTime.of(2024, 1, 3, 12, 30), "异常", "C类", 58),
                create("赵六", 35, new BigDecimal("8888.88"), 0.50,
                        LocalDate.of(2024, 1, 4), LocalDateTime.of(2024, 1, 4, 13, 30), "正常", "A类", 91)
        );
    }
}
