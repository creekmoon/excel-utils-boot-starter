package cn.creekmoon.excel.core.W.title.ext;

import lombok.Getter;

import java.util.function.Predicate;

/*条件样式*/
@Getter
public class ConditionCellStyle<R> {
    protected Predicate<R> condition;
    protected ExcelCellStyle defaultCellStyle;

    public ConditionCellStyle(ExcelCellStyle defaultCellStyle, Predicate<R> condition) {
        this.defaultCellStyle = defaultCellStyle;
        this.condition = condition;
    }


    public static <T> ConditionCellStyle<T> of(ExcelCellStyle defaultCellStyle, Predicate<T> condition) {
        return new ConditionCellStyle<T>(defaultCellStyle, condition);
    }
}
