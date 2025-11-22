package cn.creekmoon.excel.core.R.reader.cell;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.StringConverter;
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


    public HutoolCellReader(ExcelImport parent, String sheetRid, String sheetName, Supplier newObjectSupplier) {
        super(parent);
        super.readerResult = new CellReaderResult();
        super.sheetRid = sheetRid;
        super.sheetName = sheetName;
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

            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolCellReader.read] 开始读取sheet: rId={}, sheetName={}", 
                        sheetRid, sheetName);
            }
            
            Excel07SaxReader excel07SaxReader = initSaxReader(templateConsistencyCheckCells);
            /*第一个参数 文件流  第二个参数 sheetRid 直接定位到指定sheet*/
            excel07SaxReader.read(this.getParent().sourceFile.getInputStream(), getParent().rid2SheetNameBiMap.get(sheetRid));
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolCellReader.read] Sheet读取完成: rId={}, sheetName={}", 
                        sheetRid, sheetName);
            }

            /*模版一致性检查失败*/
            if (TEMPLATE_CONSISTENCY_CHECK_ENABLE && !templateConsistencyCheckCells.isEmpty()) {
                getReadResult().EXISTS_READ_FAIL.set(true);
                getReadResult().errorCount.incrementAndGet();
                getReadResult().errorReport.append(StrFormatter.format(TITLE_CHECK_ERROR));
            }

        } catch (Exception e) {
            log.error("SaxReader读取Excel文件异常", e);
        } finally {
            getReadResult().readSuccessTime = LocalDateTime.now();
            /*释放信号量*/
            ExcelUtilsConfig.importParallelSemaphore.release();
        }
        return getReadResult();
    }


    @Override
    public CellReaderResult<R> getReadResult() {
        return (CellReaderResult) readerResult;
    }


    /**
     * 初始化SAX读取器
     *
     * @param
     * @param templateConsistencyCheckCells
     * @retur
     */
    Excel07SaxReader initSaxReader(Set<String> templateConsistencyCheckCells) {
        currentNewObject = newObjectSupplier.get();


        /*返回一个Sax读取器*/
        return new Excel07SaxReader(new RowHandler() {

            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {

            }

            @Override
            public void handleCell(int sheetIndex, long rowIndex, int cellIndex, Object value, CellStyle xssfCellStyle) {
                // 由于已通过rId精准定位sheet，无需在回调中过滤
                int colIndex = cellIndex;

                /*解析单个单元格*/
                if (cell2setter.size() <= 0
                        || !cell2setter.containsKey((int) rowIndex)
                        || !cell2setter.get((int) rowIndex).containsKey(colIndex)
                ) {
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
                } catch (Exception e) {
                    getReadResult().EXISTS_READ_FAIL.set(true);
                    getReadResult().errorCount.incrementAndGet();
                    getReadResult().errorReport.append(StrFormatter.format(CONVERT_FAIL_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)))
                            .append(GlobalExceptionMsgManager.getExceptionMsg(e))
                            .append(";");
                }
            }
        });
    }

}
