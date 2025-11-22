package cn.creekmoon.excel.core.R.reader;

import cn.hutool.poi.excel.sax.handler.RowHandler;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.List;

/**
 * Sheet索引标准化的RowHandler装饰器
 * 用于解决Hutool在单sheet读取模式下sheetIndex参数不准确的问题
 * 
 * 当使用 saxReader.read(stream, N) 只读取第N个sheet时，
 * Hutool会在回调中传入sheetIndex=0（因为只处理了一个sheet）
 * 此装饰器会将其修正为实际的sheet索引
 * 
 * @author JY
 */
public class SheetIndexNormalizingRowHandler implements RowHandler {

    /**
     * 被包装的原始handler
     */
    private final RowHandler delegate;

    /**
     * 真实的目标sheet索引
     */
    private final int actualSheetIndex;

    /**
     * 是否为单sheet优化模式
     */
    private final boolean isSingleSheetMode;

    /**
     * 调试模式开关
     */
    private final boolean debugger;

    /**
     * 构造函数
     *
     * @param delegate          被包装的原始RowHandler
     * @param actualSheetIndex  真实的sheet索引
     * @param isSingleSheetMode 是否为单sheet优化模式
     * @param debugger          调试模式开关
     */
    public SheetIndexNormalizingRowHandler(RowHandler delegate, int actualSheetIndex, boolean isSingleSheetMode, boolean debugger) {
        this.delegate = delegate;
        this.actualSheetIndex = actualSheetIndex;
        this.isSingleSheetMode = isSingleSheetMode;
        this.debugger = debugger;
    }

    /**
     * 处理行数据，修正sheetIndex后转发给delegate
     */
    @Override
    public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {
        // 在单sheet模式下，Hutool传入的sheetIndex是0
        // 需要将其修正为真实的sheetIndex
        int normalizedSheetIndex = isSingleSheetMode ? actualSheetIndex : sheetIndex;
        
        if (debugger) {
            org.slf4j.LoggerFactory.getLogger(SheetIndexNormalizingRowHandler.class).info(
                "[DEBUGGER][SheetIndexNormalizingRowHandler.handle] " +
                "原始sheetIndex={}, 修正后sheetIndex={}, rowIndex={}, rowListSize={}, " +
                "isSingleSheetMode={}, actualSheetIndex={}", 
                sheetIndex, normalizedSheetIndex, rowIndex, rowList == null ? 0 : rowList.size(),
                isSingleSheetMode, actualSheetIndex
            );
        }
        
        delegate.handle(normalizedSheetIndex, rowIndex, rowList);
    }

    /**
     * 处理单元格数据，修正sheetIndex后转发给delegate
     */
    @Override
    public void handleCell(int sheetIndex, long rowIndex, int cellIndex, Object value, CellStyle xssfCellStyle) {
        // 在单sheet模式下，Hutool传入的sheetIndex是0
        // 需要将其修正为真实的sheetIndex
        int normalizedSheetIndex = isSingleSheetMode ? actualSheetIndex : sheetIndex;
        
        if (debugger) {
            org.slf4j.LoggerFactory.getLogger(SheetIndexNormalizingRowHandler.class).info(
                "[DEBUGGER][SheetIndexNormalizingRowHandler.handleCell] " +
                "原始sheetIndex={}, 修正后sheetIndex={}, rowIndex={}, cellIndex={}, " +
                "isSingleSheetMode={}, actualSheetIndex={}", 
                sheetIndex, normalizedSheetIndex, rowIndex, cellIndex,
                isSingleSheetMode, actualSheetIndex
            );
        }
        
        delegate.handleCell(normalizedSheetIndex, rowIndex, cellIndex, value, xssfCellStyle);
    }

    /**
     * 分析完成后的回调，直接转发给delegate
     */
    @Override
    public void doAfterAllAnalysed() {
        if (debugger) {
            org.slf4j.LoggerFactory.getLogger(SheetIndexNormalizingRowHandler.class).info(
                "[DEBUGGER][SheetIndexNormalizingRowHandler.doAfterAllAnalysed] 分析完成回调被触发: actualSheetIndex={}, isSingleSheetMode={}", 
                actualSheetIndex, isSingleSheetMode
            );
        }
        delegate.doAfterAllAnalysed();
    }
}

