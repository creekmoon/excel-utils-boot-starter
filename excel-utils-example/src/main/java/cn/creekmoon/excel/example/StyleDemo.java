package cn.creekmoon.excel.example;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 * 样式展示专用数据类
 */
@Data
@Builder
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
}
