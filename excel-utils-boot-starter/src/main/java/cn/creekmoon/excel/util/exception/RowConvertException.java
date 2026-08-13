package cn.creekmoon.excel.util.exception;

import cn.creekmoon.excel.util.ExcelConstants;

import java.util.List;
import java.util.stream.Collectors;

/**
 * 单行内一个或多个字段转换失败
 * <p>
 * 继承 CheckedExcelException 以复用内置的自定义异常判定：本异常只让当前行失败，不中断整表读取。
 * getMessage() 返回合并后的整行消息；字段级明细通过 getFieldFailures() 获取，用于逐列生成行错误。
 */
public class RowConvertException extends CheckedExcelException {

    private final List<FieldFailure> fieldFailures;

    /**
     * 用本行收集到的全部失败字段构造异常
     * 各字段消息按统一分隔符合并后作为整行消息
     *
     * @param fieldFailures 本行失败字段明细，由 rowConvert 逐列收集后传入，不可为空列表
     */
    public RowConvertException(List<FieldFailure> fieldFailures) {
        super(fieldFailures.stream()
                .map(FieldFailure::message)
                .collect(Collectors.joining(ExcelConstants.ROW_ERROR_MSG_SEPARATOR)));
        this.fieldFailures = fieldFailures;
    }

    /**
     * 获取本行全部失败字段的明细
     * 每个元素对应一个失败的列
     */
    public List<FieldFailure> getFieldFailures() {
        return fieldFailures;
    }

    /**
     * 字段级失败明细
     *
     * @param title   失败列的标题名
     * @param message 失败原因，未施加 [失败!] 前缀，由调用方决定展示时是否装饰
     */
    public record FieldFailure(String title, String message) {
    }
}
