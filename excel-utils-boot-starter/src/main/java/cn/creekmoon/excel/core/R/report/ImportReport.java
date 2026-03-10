package cn.creekmoon.excel.core.R.report;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel导入结果报告（结构化，支持JSON序列化）
 * 大版本升级：替代 rowIndex2msg 成为唯一事实来源
 *
 * @param <R> 导入的数据类型
 */
@Data
public class ImportReport<R> {

    /**
     * 任务唯一标识
     */
    private String taskId;

    /**
     * Sheet的rId（Excel内部关系ID）
     */
    private String sheetRid;

    /**
     * Sheet的名称
     */
    private String sheetName;

    /**
     * 标题行号（0-based）
     */
    private int titleRowIndex;

    /**
     * 配置的首行数据行号（0-based）
     */
    private int firstRowIndex;

    /**
     * 配置的末行数据行号（0-based）
     */
    private int latestRowIndex;

    /**
     * 本次实际处理到的首行数据行号（0-based，可能因fail-fast而与配置不同）
     */
    private Integer processedFirstRowIndex;

    /**
     * 本次实际处理到的末行数据行号（0-based，可能因fail-fast或阈值停止而与配置不同）
     */
    private Integer processedLastRowIndex;

    /**
     * 实际进入处理的总行数（不含标题行，不含被过滤的空白行）
     */
    private int totalRows = 0;

    /**
     * 成功转换并处理的行数
     */
    private int successRows = 0;

    /**
     * 转换或处理失败的行数
     */
    private int errorRows = 0;

    /**
     * 全局错误码（如模板不匹配、错误阈值触发、致命异常等）
     */
    private String globalErrorCode;

    /**
     * 全局错误消息
     */
    private String globalErrorMessage;

    /**
     * 行级错误明细列表
     */
    private List<RowError> rowErrors = new ArrayList<>();

    /**
     * 成功导入的数据（可选，由 ImportOptions 控制是否收集）
     */
    private List<R> data;


    /**
     * 行级错误详情
     */
    @Data
    public static class RowError {
        /**
         * 行号（0-based）
         */
        private int rowIndex;

        /**
         * 错误码（如 TEMPLATE_MISMATCH / VALIDATION_ERROR / FATAL）
         */
        private String code;

        /**
         * 错误消息
         */
        private String message;

        /**
         * 字段名（可选，标题名或单元格坐标）
         */
        private String field;

        /**
         * 原始值（可选，默认不采集以控制内存）
         */
        private String raw;

        public RowError(int rowIndex, String code, String message) {
            this.rowIndex = rowIndex;
            this.code = code;
            this.message = message;
        }

        public RowError(int rowIndex, String code, String message, String field) {
            this.rowIndex = rowIndex;
            this.code = code;
            this.message = message;
            this.field = field;
        }
    }
}
