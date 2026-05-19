package cn.creekmoon.excel.core.W.title;

import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.core.W.Writer;
import cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle;
import cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle;
import cn.creekmoon.excel.core.W.title.ext.Title;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.handler.context.SheetWriteHandlerContext;
import com.alibaba.excel.write.metadata.WriteSheet;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.*;
import java.util.stream.Collectors;

/**
 * 基于 EasyExcel 4.x 的流式写入器，替换 HutoolTitleWriter。
 * 支持多级表头、条件样式、reset 分段写、多 sheet 共享同一 ExcelWriter。
 */
@Slf4j
public class EasyExcelTitleWriter<R> extends TitleWriter<R> {

    protected ExcelExport parent;
    protected ExcelWriter excelWriter;
    protected WriteSheet writeSheet;
    protected int currentRow = 0;

    /* 用于条件样式：行号 → 数据对象的映射 */
    private final Map<Integer, R> rowIndex2Data = new HashMap<>();
    private static final int BATCH_SIZE = 100;
    private static final int DEFAULT_COLUMN_WIDTH_COUNT = 500;
    private static final int DEFAULT_COLUMN_WIDTH = 20;

    private String sheetInitKey(int sheetIndex) {
        return "SHEET_INIT_" + sheetIndex;
    }

    public EasyExcelTitleWriter(ExcelExport parent, Integer sheetIndex) {
        this.parent = parent;
        this.sheetIndex = sheetIndex;
        this.sheetName = "Sheet" + sheetIndex;
    }

    public EasyExcelTitleWriter(ExcelExport parent, Integer sheetIndex, String sheetName) {
        this.parent = parent;
        this.sheetIndex = sheetIndex;
        this.sheetName = sheetName;
    }

    @Override
    public <T> EasyExcelTitleWriter<T> reset(Class<T> newDataClass) {
        return reset(newDataClass, 0);
    }

    @Override
    public <T> EasyExcelTitleWriter<T> reset(Class<T> newDataClass, int rowIndexOffset) {
        EasyExcelTitleWriter<T> newWriter = new EasyExcelTitleWriter<>(this.parent, this.sheetIndex, this.sheetName);
        newWriter.excelWriter = this.excelWriter;
        newWriter.currentRow = this.currentRow + rowIndexOffset;
        newWriter.titleRowIndex = newWriter.currentRow;
        newWriter.firstRowIndex = null;
        newWriter.latestRowIndex = null;
        return newWriter;
    }

    @Override
    public int countTitles() {
        return titles.size();
    }

    @Override
    public Workbook getWorkbook() {
        return getExcelWriter().writeContext().writeWorkbookHolder().getWorkbook();
    }

    /**
     * 获取或创建 ExcelWriter。多 sheet 共享同一个实例。
     */
    private ExcelWriter getExcelWriter() {
        if (excelWriter == null && parent.sheetIndex2SheetWriter.containsKey(sheetIndex - 1)) {
            Writer prevWriter = parent.sheetIndex2SheetWriter.get(sheetIndex - 1);
            if (prevWriter instanceof EasyExcelTitleWriter<?> prev) {
                excelWriter = prev.excelWriter;
            }
        }
        if (excelWriter == null) {
            excelWriter = EasyExcel.write(parent.getResultFilePath())
                    .registerWriteHandler(new ColumnWidthSheetWriteHandler())
                    .registerWriteHandler(new ConditionStyleCellWriteHandler())
                    .build();
        }
        return excelWriter;
    }

    @Override
    public EasyExcelTitleWriter<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex) {
        this.titleRowIndex = titleRowIndex;
        this.firstRowIndex = firstDataRowIndex;
        this.latestRowIndex = lastDataRowIndex;
        return this;
    }

    @Override
    public EasyExcelTitleWriter<R> range(int startRowIndex, int lastRowIndex) {
        this.titleRowIndex = startRowIndex;
        this.firstRowIndex = null;
        this.latestRowIndex = lastRowIndex;
        return this;
    }

    @Override
    public EasyExcelTitleWriter<R> range(int startRowIndex) {
        this.titleRowIndex = startRowIndex;
        this.firstRowIndex = null;
        this.latestRowIndex = null;
        return this;
    }

    @Override
    protected void preWrite() {
        // sheet 初始化顺序检查
        if (sheetIndex != null && sheetIndex > 0) {
            if (!parent.metadatas.containsKey(sheetInitKey(sheetIndex - 1))) {
                throw new IllegalStateException("切换sheet页失败：sheetIndex=" + sheetIndex
                        + " 的首次写入要求 sheetIndex=" + (sheetIndex - 1)
                        + " 已经执行过一次 write() 完成初始化（允许 write(emptyList) 仅写表头）。");
            }
        }

        if (writeSheet == null) {
            ExcelWriterSheetBuilder builder = EasyExcel.writerSheet(sheetIndex, sheetName);

            if (!isTitleInitialized()) {
                initTitleConfig();
                List<List<String>> head = buildHead();
                builder.head(head);
                if (titleRowIndex != null) {
                    builder.relativeHeadRowIndex(titleRowIndex);
                }
            }

            writeSheet = builder.build();
        }

        parent.metadatas.put(sheetInitKey(sheetIndex), Boolean.TRUE);
    }

    /**
     * 初始化表头深度、首行/末行等配置。
     */
    private void initTitleConfig() {
        if (titleRowIndex == null) {
            titleRowIndex = 0;
        }
        if (latestRowIndex == null) {
            latestRowIndex = Integer.MAX_VALUE;
        }

        MAX_TITLE_DEPTH = titles.stream()
                .map(x -> countSeparator(x.titleName, x.PARENT_TITLE_SEPARATOR) + 1)
                .max(Comparator.naturalOrder())
                .orElse(1);

        if (firstRowIndex == null) {
            firstRowIndex = titleRowIndex + MAX_TITLE_DEPTH;
        }
        if (currentRow < firstRowIndex) {
            currentRow = firstRowIndex;
        }
    }

    private int countSeparator(String str, String separator) {
        int count = 0;
        int idx = 0;
        while ((idx = str.indexOf(separator, idx)) != -1) {
            count++;
            idx += separator.length();
        }
        return count;
    }

    /**
     * 将 Title 列表转换为 EasyExcel 多级表头格式 List<List<String>>。
     * 较浅的标题会自动补齐到最大深度，以匹配 Hutool 的手动合并行为。
     */
    private List<List<String>> buildHead() {
        List<List<String>> head = new ArrayList<>();
        for (Title title : titles) {
            String[] parts = title.titleName.split(title.PARENT_TITLE_SEPARATOR);
            List<String> headColumn = new ArrayList<>(Arrays.asList(parts));

            // 补齐到 MAX_TITLE_DEPTH
            String lastPart = parts[parts.length - 1];
            while (headColumn.size() < MAX_TITLE_DEPTH) {
                headColumn.add(lastPart);
            }
            head.add(headColumn);
        }
        return head;
    }

    @Override
    protected void doWrite(List<R> data) {
        int maxRows = latestRowIndex - currentRow + 1;
        List<R> dataToWrite = data;
        if (maxRows < data.size() && latestRowIndex != Integer.MAX_VALUE) {
            dataToWrite = data.subList(0, Math.max(0, maxRows));
        }

        List<List<R>> batches = partition(dataToWrite, BATCH_SIZE);
        if (batches.isEmpty()) {
            // 空数据时仍触发 sheet 创建（确保仅写表头的 sheet 能被 stopWrite 重排序）
            getExcelWriter().write(Collections.emptyList(), writeSheet);
        }
        for (List<R> batch : batches) {
            // 记录行号映射，供条件样式 handler 使用
            int batchStartRow = currentRow;
            for (int i = 0; i < batch.size(); i++) {
                rowIndex2Data.put(batchStartRow + i, batch.get(i));
            }

            List<List<Object>> rows = batch.stream()
                    .map(this::changeToCellValues)
                    .collect(Collectors.toList());

            getExcelWriter().write(rows, writeSheet);
            currentRow += batch.size();
        }
    }

    private List<List<R>> partition(List<R> list, int size) {
        List<List<R>> result = new ArrayList<>();
        for (int i = 0; i < list.size(); i += size) {
            result.add(list.subList(i, Math.min(i + size, list.size())));
        }
        return result;
    }

    private List<Object> changeToCellValues(R dataObject) {
        List<Object> row = new ArrayList<>();
        for (Title title : titles) {
            Object data = null;
            try {
                data = title.valueFunction.apply(dataObject);
            } catch (Exception e) {
                log.debug("[EasyExcelTitleWriter] 生成Excel获取数据值时发生错误!已经忽略错误并设置为NULL值!", e);
            }
            row.add(data);
        }
        return row;
    }

    @Override
    protected CellStyle getRunningTimeCellStyle(ExcelCellStyle style) {
        Object cached = parent.metadatas.get(style);
        if (cached != null) {
            return (CellStyle) cached;
        }
        Workbook workbook = getWorkbook();
        CellStyle newCellStyle = workbook.createCellStyle();
        style.getStyleInitializer().accept(workbook, newCellStyle);
        parent.metadatas.put(style, newCellStyle);
        return newCellStyle;
    }

    @Override
    protected void onSwitchSheet() {
        // EasyExcel 通过 sheet index 自动处理
    }

    @Override
    protected void stopWrite() {
        // 计数：只有最后一个 writer 才真正 finish ExcelWriter
        int writerCount = parent.sheetIndex2SheetWriter.size();
        int closedCount = (Integer) parent.metadatas.getOrDefault("EASYEXCEL_WRITER_CLOSE_COUNT", 0) + 1;
        parent.metadatas.put("EASYEXCEL_WRITER_CLOSE_COUNT", closedCount);

        if (closedCount >= writerCount && excelWriter != null) {
            excelWriter.finish();
        }
    }

    private boolean isTitleInitialized() {
        return this.MAX_TITLE_DEPTH != null;
    }

    /**
     * 条件样式 CellWriteHandler：在单元格写入完成后，根据数据和标题配置回写 POI CellStyle。
     */
    private class ConditionStyleCellWriteHandler implements CellWriteHandler {

        @Override
        public int order() {
            return 60000;
        }

        @Override
        public void afterCellDispose(CellWriteHandlerContext context) {
            if (Boolean.TRUE.equals(context.getHead())) {
                return;
            }

            String currentSheetName = context.getWriteSheetHolder().getSheetName();

            /* 多 sheet 共享同一个 ExcelWriter，需根据 sheetName 定位到对应的 writer 实例 */
            EasyExcelTitleWriter<?> targetWriter = null;
            for (Writer w : parent.sheetIndex2SheetWriter.values()) {
                if (w instanceof EasyExcelTitleWriter<?> etw && currentSheetName.equals(etw.sheetName)) {
                    targetWriter = etw;
                    break;
                }
            }
            if (targetWriter == null) {
                return;
            }

            int rowIndex = context.getCell().getRowIndex();
            int colIndex = context.getCell().getColumnIndex();

            if (colIndex < 0 || colIndex >= targetWriter.titles.size()) {
                return;
            }

            Object data = targetWriter.rowIndex2Data.get(rowIndex);
            if (data == null) {
                return;
            }

            Title title = targetWriter.titles.get(colIndex);
            for (ConditionCellStyle conditionStyle : title.conditionCellStyleList) {
                if (conditionStyle == null || conditionStyle.getCondition() == null) {
                    continue;
                }
                if (conditionStyle.getCondition().test(data)) {
                    CellStyle poiStyle = targetWriter.getRunningTimeCellStyle(conditionStyle.getDefaultCellStyle());
                    context.getCell().setCellStyle(poiStyle);
                }
            }
        }
    }

    /**
     * 列宽设置 SheetWriteHandler：在 sheet 创建后设置默认列宽。
     */
    private class ColumnWidthSheetWriteHandler implements SheetWriteHandler {
        @Override
        public void afterSheetCreate(SheetWriteHandlerContext context) {
            String currentSheetName = context.getWriteSheetHolder().getSheetName();
            if (!currentSheetName.equals(EasyExcelTitleWriter.this.sheetName)) {
                return;
            }
            org.apache.poi.ss.usermodel.Sheet sheet = context.getWriteSheetHolder().getSheet();
            for (int i = 0; i < DEFAULT_COLUMN_WIDTH_COUNT; i++) {
                try {
                    sheet.setColumnWidth(i, DEFAULT_COLUMN_WIDTH * 256);
                } catch (Exception ignored) {
                }
            }
        }
    }
}
