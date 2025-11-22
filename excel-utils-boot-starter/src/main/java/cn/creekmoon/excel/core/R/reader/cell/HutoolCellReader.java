package cn.creekmoon.excel.core.R.reader.cell;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.StringConverter;
import cn.creekmoon.excel.core.R.reader.SheetIndexNormalizingRowHandler;
import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.core.R.readerResult.cell.CellReaderResult;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.creekmoon.excel.util.ExcelCellUtils;
import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;
import cn.creekmoon.excel.util.exception.GlobalExceptionMsgManager;
import cn.hutool.core.text.StrFormatter;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

import static cn.creekmoon.excel.util.ExcelConstants.*;

@Slf4j
public class HutoolCellReader<R> extends CellReader<R> {


    public HutoolCellReader(ExcelImport parent, Integer sheetIndex, Supplier newObjectSupplier) {
        super(parent);
        super.readerResult = new CellReaderResult();
        super.sheetIndex = sheetIndex;
        super.newObjectSupplier = newObjectSupplier;
    }

    @Override
    public <T> HutoolCellReader<R> addConvert(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter) {

        return addConvert(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), convert, setter);
    }

    @Override
    public HutoolCellReader<R> addConvert(String cellReference, BiConsumer<R, String> reader) {
        return addConvert(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), x -> x, reader);
    }

    @Override
    public <T> HutoolCellReader<R> addConvert(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        cell2converts.computeIfAbsent(rowIndex, HashMap::new);
        cell2converts.get(rowIndex).put(colIndex, convert);
        cell2setter.computeIfAbsent(rowIndex, HashMap::new);
        cell2setter.get(rowIndex).put(colIndex, setter);
        return this;
    }

    @Override
    public HutoolCellReader<R> addConvert(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvert(rowIndex, colIndex, x -> x, setter);
    }


    @Override
    public <T> HutoolCellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvertAndSkipEmpty(rowIndex, colIndex, x -> x, setter);
    }

    @Override
    public <T> HutoolCellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        skipEmptyCells.computeIfAbsent(rowIndex, HashSet::new);
        skipEmptyCells.get(rowIndex).add(colIndex);
        return addConvert(rowIndex, colIndex, convert, setter);
    }

    @Override
    public <T> HutoolCellReader<R> addConvertAndSkipEmpty(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        return addConvertAndSkipEmpty(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), convert, setter);
    }

    @Override
    public HutoolCellReader<R> addConvertAndSkipEmpty(String cellReference, BiConsumer<R, String> setter) {
        return addConvertAndSkipEmpty(cellReference, x -> x, setter);
    }

    @Override
    public <T> HutoolCellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        mustExistCells.computeIfAbsent(rowIndex, HashSet::new);
        mustExistCells.get(rowIndex).add(colIndex);
        return addConvert(rowIndex, colIndex, convert, setter);
    }


    @Override
    public HutoolCellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvertAndMustExist(rowIndex, colIndex, x -> x, setter);
    }

    @Override
    public HutoolCellReader<R> addConvertAndMustExist(String cellReference, BiConsumer<R, String> setter) {
        return addConvertAndMustExist(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), setter);
    }

    @Override
    public ReaderResult<R> read(ExConsumer<R> consumer) throws Exception {
        return read().consume(consumer);
    }


    @Override
    public ReaderResult<R> read() throws InterruptedException, IOException {
        if (getParent().debugger) {
            log.info("[DEBUGGER][HutoolCellReader.read] 开始读取: 目标sheetIndex={}, 文件名={}", 
                sheetIndex, getParent().sourceFile.getOriginalFilename());
        }

        //新版读取 使用SAX读取模式
        ExcelUtilsConfig.importParallelSemaphore.acquire();
        getReadResult().readStartTime = LocalDateTime.now();
        try {
            /*模版一致性检查:  获取声明的所有CELL, 接下来如果读取到cell就会移除, 当所有cell命中时说明单元格是一致的.*/
            Set<String> templateConsistencyCheckCells = new HashSet<>();
            if (TEMPLATE_CONSISTENCY_CHECK_ENABLE) {
                cell2setter.forEach((rowIndex, colIndexMap) -> {
                    colIndexMap.forEach((colIndex, var) -> {
                        templateConsistencyCheckCells.add(ExcelCellUtils.excelIndexToCell(rowIndex, colIndex));
                    });
                });
            }

            // 获取sheetIndex对应的rid
            String rid = getParent().getSheetRid(sheetIndex);
            
            // 确定是否使用单sheet优化模式
            boolean isSingleSheetMode;
            if (rid == null || rid.isEmpty()) {
                // 映射失败，降级使用-1（读取所有sheet）
                rid = "-1";
                isSingleSheetMode = false;
                if (getParent().debugger) {
                    log.warn("[DEBUGGER][HutoolCellReader.read] 无法获取rid，降级使用-1: sheetIndex={}", sheetIndex);
                }
            } else {
                // 使用单sheet优化
                isSingleSheetMode = true;
                if (getParent().debugger) {
                    log.info("[DEBUGGER][HutoolCellReader.read] 使用单sheet优化: sheetIndex={} → rid={}", 
                        sheetIndex, rid);
                }
            }

            /*创建原始RowHandler*/
            RowHandler originalHandler = createRowHandler(templateConsistencyCheckCells);
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolCellReader.read] 创建原始RowHandler成功");
            }
            
            /*用适配器包装*/
            RowHandler wrappedHandler = new SheetIndexNormalizingRowHandler(
                originalHandler,
                sheetIndex,
                isSingleSheetMode,
                getParent().debugger
            );
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolCellReader.read] 创建适配器包装成功: actualSheetIndex={}", sheetIndex);
            }
            
            /*创建SaxReader*/
            Excel07SaxReader excel07SaxReader = new Excel07SaxReader(wrappedHandler);
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolCellReader.read] 准备调用saxReader.read(): rid={}", rid);
            }
            
            /*使用rid读取（可能是单sheet或所有sheet）*/
            excel07SaxReader.read(this.getParent().sourceFile.getInputStream(), rid);
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolCellReader.read] saxReader.read()执行完成");
            }

            /*模版一致性检查失败*/
            if (TEMPLATE_CONSISTENCY_CHECK_ENABLE && !templateConsistencyCheckCells.isEmpty()) {
                getReadResult().EXISTS_READ_FAIL.set(true);
                getReadResult().errorCount.incrementAndGet();
                getReadResult().errorReport.append(StrFormatter.format(TITLE_CHECK_ERROR));
            }

        } catch (Exception e) {
            log.error("SaxReader读取Excel文件异常", e);
            if (getParent().debugger) {
                log.error("[DEBUGGER][HutoolCellReader.read] 读取异常: {}", e.getMessage(), e);
            }
        } finally {
            getReadResult().readSuccessTime = LocalDateTime.now();
            /*释放信号量*/
            ExcelUtilsConfig.importParallelSemaphore.release();
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolCellReader.read] 读取完成: 耗时={}ms", 
                    java.time.Duration.between(getReadResult().readStartTime, getReadResult().readSuccessTime).toMillis());
            }
        }
        return getReadResult();
    }


    @Override
    public CellReaderResult<R> getReadResult() {
        return (CellReaderResult) readerResult;
    }


    /**
     * 创建RowHandler（业务处理逻辑）
     *
     * @param templateConsistencyCheckCells 模板一致性检查单元格集合
     * @return RowHandler实例
     */
    RowHandler createRowHandler(Set<String> templateConsistencyCheckCells) {
        Integer targetSheetIndex = sheetIndex;
        currentNewObject = newObjectSupplier.get();

        return new RowHandler() {
            int currentSheetIndex = 0;

            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {

            }

            @Override
            public void handleCell(int sheetIndex, long rowIndex, int cellIndex, Object value, CellStyle xssfCellStyle) {
                currentSheetIndex = sheetIndex;

                int colIndex = cellIndex;

                if (getParent().debugger) {
                    log.info("[DEBUGGER][HutoolCellReader.handleCell] 接收回调: sheetIndex={}, rowIndex={}, cellIndex={}, value={}, targetSheetIndex={}", 
                        sheetIndex, rowIndex, cellIndex, value, targetSheetIndex);
                }

                if (targetSheetIndex != currentSheetIndex) {
                    if (getParent().debugger) {
                        log.info("[DEBUGGER][HutoolCellReader.handleCell] Sheet过滤: targetSheetIndex({}) != currentSheetIndex({}), 跳过", 
                            targetSheetIndex, currentSheetIndex);
                    }
                    return;
                }

                /*解析单个单元格*/
                if (cell2setter.size() <= 0
                        || !cell2setter.containsKey((int) rowIndex)
                        || !cell2setter.get((int) rowIndex).containsKey(colIndex)
                ) {
                    if (getParent().debugger) {
                        log.info("[DEBUGGER][HutoolCellReader.handleCell] 单元格不在配置中: rowIndex={}, colIndex={}, 跳过", 
                            rowIndex, colIndex);
                    }
                    return;
                }
                /*标题一致性检查*/
                templateConsistencyCheckCells.remove(ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex));

                try {
                    ExFunction cellConverter = cell2converts.get((int) rowIndex).get(colIndex);
                    BiConsumer cellConsumer = cell2setter.get((int) rowIndex).get(colIndex);
                    String cellValue = StringConverter.parse(value);
                    /*检查必填项/检查可填项*/
                    if (StrUtil.isBlank(cellValue)) {
                        if (mustExistCells.containsKey((int) rowIndex)
                                && mustExistCells.get((int) rowIndex).contains(colIndex)) {
                            throw new CheckedExcelException(StrFormatter.format(FIELD_LACK_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)));
                        }
                        if (skipEmptyCells.containsKey((int) rowIndex)
                                && skipEmptyCells.get((int) rowIndex).contains(colIndex)) {
                            return;
                        }
                    }
                    Object apply = cellConverter.apply(cellValue);
                    cellConsumer.accept(currentNewObject, apply);
                    getReadResult().setData((R) currentNewObject);
                    
                    if (getParent().debugger) {
                        log.info("[DEBUGGER][HutoolCellReader.handleCell] 单元格处理成功: rowIndex={}, colIndex={}, cellValue={}", 
                            rowIndex, colIndex, cellValue);
                    }
                } catch (Exception e) {
                    getReadResult().EXISTS_READ_FAIL.set(true);
                    getReadResult().errorCount.incrementAndGet();
                    getReadResult().errorReport.append(StrFormatter.format(CONVERT_FAIL_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)))
                            .append(GlobalExceptionMsgManager.getExceptionMsg(e))
                            .append(";");
                    
                    if (getParent().debugger) {
                        log.error("[DEBUGGER][HutoolCellReader.handleCell] 单元格处理异常: rowIndex={}, colIndex={}, error={}", 
                            rowIndex, colIndex, e.getMessage(), e);
                    }
                }
            }
        };
    }

    /**
     * 初始化SAX读取器（保留用于向后兼容）
     *
     * @param templateConsistencyCheckCells 模板一致性检查单元格集合
     * @return Excel07SaxReader实例
     */
    @Deprecated
    Excel07SaxReader initSaxReader(Set<String> templateConsistencyCheckCells) {
        return new Excel07SaxReader(createRowHandler(templateConsistencyCheckCells));
    }

    @Override
    public Integer getSheetIndex() {
        return getParent().sheetIndex2ReaderBiMap.getKey(this);
    }
}
