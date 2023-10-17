package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.converter.StringConverter;
import cn.creekmoon.excelUtils.exception.CheckedExcelException;
import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;
import cn.creekmoon.excelUtils.util.ExcelCellUtils;
import cn.hutool.core.text.StrFormatter;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.BiConsumer;

import static cn.creekmoon.excelUtils.core.ExcelConstants.CONVERT_FAIL_MSG;
import static cn.creekmoon.excelUtils.core.ExcelConstants.FIELD_LACK_MSG;

@Slf4j
public class CellReader<R> {

    public SheetReaderContext sheetReaderContext;

    protected ExcelImport parent;


    /**
     * 获取SHEET页的总行数
     *
     * @return
     */
    @SneakyThrows
    public Long getSheetRowCount() {
        return parent.getSheetRowCount(sheetReaderContext.sheetIndex);
    }

    public <T> CellReader<R> addConvert(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter) {

        return addConvert(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), convert, setter);
    }

    public CellReader<R> addConvert(String cellReference, BiConsumer<R, String> reader) {
        return addConvert(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), x -> x, reader);
    }

    public <T> CellReader<R> addConvert(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        sheetReaderContext.cell2converts.computeIfAbsent(rowIndex, HashMap::new);
        sheetReaderContext.cell2converts.get(rowIndex).put(colIndex, convert);
        sheetReaderContext.cell2consumers.computeIfAbsent(rowIndex, HashMap::new);
        sheetReaderContext.cell2consumers.get(rowIndex).put(colIndex, setter);
        return this;
    }

    public <T> CellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvertAndSkipEmpty(rowIndex, colIndex, x -> x, setter);
    }

    public <T> CellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        sheetReaderContext.skipEmptyCells.computeIfAbsent(rowIndex, HashSet::new);
        sheetReaderContext.skipEmptyCells.get(rowIndex).add(colIndex);
        return addConvert(rowIndex, colIndex, convert, setter);
    }

    public <T> CellReader<R> addConvertAndSkipEmpty(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        return addConvertAndSkipEmpty(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), convert, setter);
    }

    public <T> CellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        sheetReaderContext.mustExistCells.computeIfAbsent(rowIndex, HashSet::new);
        sheetReaderContext.mustExistCells.get(rowIndex).add(colIndex);
        return addConvert(rowIndex, colIndex, convert, setter);
    }

    public <T> CellReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        return addConvertAndMustExist(ExcelCellUtils.excelCellToRowIndex(title), ExcelCellUtils.excelCellToColumnIndex(title), convert, setter);
    }

    public CellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvertAndMustExist(rowIndex, colIndex, x -> x, setter);
    }

    public CellReader<R> addConvertAndMustExist(String cellReference, BiConsumer<R, String> setter) {
        return addConvertAndMustExist(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), setter);
    }


    /**
     * 添加校验阶段后置处理器 当所有的convert执行完成后会执行这个操作做最后的校验处理
     *
     * @param postProcessor 后置处理器
     * @param <T>
     * @return
     */
    public <T> CellReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor) {
        if (postProcessor != null) {
            this.sheetReaderContext.cellConvertPostProcessors.add(postProcessor);
        }
        return this;
    }

    public void read(ExConsumer<R> consumer) throws CheckedExcelException, IOException {

        //新版读取 使用SAX读取模式
        Excel07SaxReader excel07SaxReader = initSaxReader(sheetReaderContext.sheetIndex, consumer);
        /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
        excel07SaxReader.read(this.parent.file.getInputStream(), -1);
        if (sheetReaderContext.errorReport.length() > 0) {
            throw new CheckedExcelException(sheetReaderContext.errorReport.toString());
        }


    }

    public R read() throws CheckedExcelException, IOException {
        AtomicReference<R> result = new AtomicReference<>();
        read(result::set);
        return result.get();
    }

    /**
     * 初始化SAX读取器
     *
     * @param targetSheetIndex 读取的sheetIndex下标
     * @param consumer
     * @return
     */
    Excel07SaxReader initSaxReader(int targetSheetIndex, ExConsumer<R> consumer) {

        sheetReaderContext.currentNewObject = sheetReaderContext.newObjectSupplier.get();

        /*返回一个Sax读取器*/
        return new Excel07SaxReader(new RowHandler() {
            int currentSheetIndex = 0;

            @Override
            public void doAfterAllAnalysed() {

                if (targetSheetIndex != currentSheetIndex) {
                    return;
                }

                /*sheet读取结束时*/
                for (ExConsumer convertPostProcessor : sheetReaderContext.cellConvertPostProcessors) {
                    if (sheetReaderContext.errorReport.length() > 0) {
                        throw new RuntimeException("导入失败!");
                    }
                    try {
                        convertPostProcessor.accept(sheetReaderContext.currentNewObject);
                    } catch (Exception e) {
                        parent.getErrorCount().incrementAndGet();
                        sheetReaderContext.errorReport.append(GlobalExceptionManager.getExceptionMsg(e));
                    }
                }

                try {
                    consumer.accept((R) sheetReaderContext.currentNewObject);
                } catch (Exception e) {
                    parent.getErrorCount().incrementAndGet();
                    sheetReaderContext.errorReport.append(GlobalExceptionManager.getExceptionMsg(e));
                }
            }


            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {
//                if (targetSheetIndex != sheetIndex) {
//                    return;
//                }
//
//                /*解析单个单元格*/
//                if (sheetReaderContext.cell2consumers.size() <= 0
//                        || !sheetReaderContext.cell2consumers.containsKey((int) rowIndex)
//                ) {
//                    return;
//                }
//
//                HashMap<Integer, BiConsumer> col2Consumer = sheetReaderContext.cell2consumers.get((int) rowIndex);
//                for (Integer colIndex : col2Consumer.keySet()) {
//                    try {
//                        ExFunction cellConverter = sheetReaderContext.cell2converts.get((int) rowIndex).get(colIndex);
//                        BiConsumer cellConsumer = sheetReaderContext.cell2consumers.get((int) rowIndex).get(colIndex);
//                        String cellValue = StringConverter.parse(rowList.get(colIndex));
//                        /*检查必填项/检查可填项*/
//                        if (StrUtil.isBlank(cellValue)) {
//                            if (sheetReaderContext.mustExistCells.containsKey((int) rowIndex)
//                                    && sheetReaderContext.mustExistCells.get((int) rowIndex).contains(colIndex)) {
//                                throw new CheckedExcelException(StrFormatter.format(FIELD_LACK_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)));
//                            }
//                            if (sheetReaderContext.skipEmptyCells.containsKey((int) rowIndex)
//                                    && sheetReaderContext.skipEmptyCells.get((int) rowIndex).contains(colIndex)) {
//                                continue;
//                            }
//                        }
//                        Object apply = cellConverter.apply(cellValue);
//                        cellConsumer.accept(sheetReaderContext.currentNewObject, apply);
//                    } catch (Exception e) {
//                        parent.getErrorCount().incrementAndGet();
//                        sheetReaderContext.errorReport.append(StrFormatter.format(CONVERT_FAIL_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)))
//                                .append(GlobalExceptionManager.getExceptionMsg(e))
//                                .append(";");
//                    }
//                }
            }

            @Override
            public void handleCell(int sheetIndex, long rowIndex, int cellIndex, Object value, CellStyle xssfCellStyle) {
                currentSheetIndex = sheetIndex;

                int colIndex = cellIndex;

                if (targetSheetIndex != currentSheetIndex) {
                    return;
                }

                /*解析单个单元格*/
                if (sheetReaderContext.cell2consumers.size() <= 0
                        || !sheetReaderContext.cell2consumers.containsKey((int) rowIndex)
                        || !sheetReaderContext.cell2consumers.get((int) rowIndex).containsKey(colIndex)
                ) {
                    return;
                }

                try {
                    ExFunction cellConverter = sheetReaderContext.cell2converts.get((int) rowIndex).get(colIndex);
                    BiConsumer cellConsumer = sheetReaderContext.cell2consumers.get((int) rowIndex).get(colIndex);
                    String cellValue = StringConverter.parse(value);
                    /*检查必填项/检查可填项*/
                    if (StrUtil.isBlank(cellValue)) {
                        if (sheetReaderContext.mustExistCells.containsKey((int) rowIndex)
                                && sheetReaderContext.mustExistCells.get((int) rowIndex).contains(colIndex)) {
                            throw new CheckedExcelException(StrFormatter.format(FIELD_LACK_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)));
                        }
                        if (sheetReaderContext.skipEmptyCells.containsKey((int) rowIndex)
                                && sheetReaderContext.skipEmptyCells.get((int) rowIndex).contains(colIndex)) {
                            return;
                        }
                    }
                    Object apply = cellConverter.apply(cellValue);
                    cellConsumer.accept(sheetReaderContext.currentNewObject, apply);
                } catch (Exception e) {
                    parent.getErrorCount().incrementAndGet();
                    sheetReaderContext.errorReport.append(StrFormatter.format(CONVERT_FAIL_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)))
                            .append(GlobalExceptionManager.getExceptionMsg(e))
                            .append(";");
                }
            }
        });
    }


}
