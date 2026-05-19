package cn.creekmoon.excel.core.R.reader.cell;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.StringConverter;
import cn.creekmoon.excel.util.ExcelCellUtils;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

/**
 * 基于 EasyExcel 4.0.3 的 SAX 流式单元格读取器
 * 替换 HutoolCellReader，保持完全一致的外部行为
 */
@Slf4j
public class EasyExcelCellReader<R> extends CellReader<R> {

    public EasyExcelCellReader(ExcelImport parent, String sheetRid, String sheetName, Supplier newObjectSupplier) {
        super(parent);
        super.sheetRid = sheetRid;
        super.sheetName = sheetName;
        super.newObjectSupplier = newObjectSupplier;
    }

    @Override
    public <T> EasyExcelCellReader<R> addConvert(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        return addConvert(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), convert, setter);
    }

    @Override
    public EasyExcelCellReader<R> addConvert(String cellReference, BiConsumer<R, String> reader) {
        return addConvert(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), x -> x, reader);
    }

    @Override
    public <T> EasyExcelCellReader<R> addConvert(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        cell2converts.computeIfAbsent(rowIndex, HashMap::new);
        cell2converts.get(rowIndex).put(colIndex, convert);
        cell2setter.computeIfAbsent(rowIndex, HashMap::new);
        cell2setter.get(rowIndex).put(colIndex, setter);
        return this;
    }

    @Override
    public EasyExcelCellReader<R> addConvert(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvert(rowIndex, colIndex, x -> x, setter);
    }

    @Override
    public <T> EasyExcelCellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvertAndSkipEmpty(rowIndex, colIndex, x -> x, setter);
    }

    @Override
    public <T> EasyExcelCellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        skipEmptyCells.computeIfAbsent(rowIndex, HashSet::new);
        skipEmptyCells.get(rowIndex).add(colIndex);
        return addConvert(rowIndex, colIndex, convert, setter);
    }

    @Override
    public <T> EasyExcelCellReader<R> addConvertAndSkipEmpty(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        return addConvertAndSkipEmpty(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), convert, setter);
    }

    @Override
    public EasyExcelCellReader<R> addConvertAndSkipEmpty(String cellReference, BiConsumer<R, String> setter) {
        return addConvertAndSkipEmpty(cellReference, x -> x, setter);
    }

    @Override
    public <T> EasyExcelCellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        mustExistCells.computeIfAbsent(rowIndex, HashSet::new);
        mustExistCells.get(rowIndex).add(colIndex);
        return addConvert(rowIndex, colIndex, convert, setter);
    }

    @Override
    public EasyExcelCellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvertAndMustExist(rowIndex, colIndex, x -> x, setter);
    }

    @Override
    public EasyExcelCellReader<R> addConvertAndMustExist(String cellReference, BiConsumer<R, String> setter) {
        return addConvertAndMustExist(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), setter);
    }

    @Override
    public R read(ExConsumer<R> consumer) throws Exception {
        R result = read();
        if (result != null && consumer != null) {
            consumer.accept(result);
        }
        return result;
    }

    @Override
    public R read() throws InterruptedException, IOException {
        ExcelUtilsConfig.importParallelSemaphore.acquire();
        try {
            /* 模板一致性检查：获取声明的所有 CELL，读取到命中后移除 */
            Set<String> templateConsistencyCheckCells = new HashSet<>();
            log.debug("[EasyExcelCellReader] 开始读取sheet: rId={}, sheetName={}",
                    sheetRid, sheetName);

            currentNewObject = newObjectSupplier.get();

            /* 初始化所有声明的单元格到一致性检查集合中 */
            for (Map.Entry<Integer, HashMap<Integer, BiConsumer>> rowEntry : cell2setter.entrySet()) {
                int rowIndex = rowEntry.getKey();
                for (Integer colIndex : rowEntry.getValue().keySet()) {
                    templateConsistencyCheckCells.add(ExcelCellUtils.excelIndexToCell(rowIndex, colIndex));
                }
            }

            EasyExcel.read(this.getParent().sourceFile.getInputStream(), new AnalysisEventListener<Map<Integer, Object>>() {
                @Override
                public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
                    // 覆盖默认行为：将 head 行也作为数据行处理
                    int rowIndex = context.readRowHolder().getRowIndex();
                    if (cell2setter.size() <= 0 || !cell2setter.containsKey(rowIndex)) {
                        return;
                    }
                    for (Map.Entry<Integer, String> entry : headMap.entrySet()) {
                        int colIndex = entry.getKey();
                        String value = entry.getValue();
                        if (!cell2setter.get(rowIndex).containsKey(colIndex)) {
                            continue;
                        }
                        templateConsistencyCheckCells.remove(ExcelCellUtils.excelIndexToCell(rowIndex, colIndex));
                        try {
                            ExFunction cellConverter = cell2converts.get(rowIndex).get(colIndex);
                            BiConsumer cellConsumer = cell2setter.get(rowIndex).get(colIndex);
                            String cellValue = StringConverter.parse(value);
                            if (cellValue == null || cellValue.isEmpty()) {
                                if (mustExistCells.containsKey(rowIndex)
                                        && mustExistCells.get(rowIndex).contains(colIndex)) {
                                    throw new CheckedExcelException(
                                            String.format(ExcelConstants.FIELD_LACK_MSG, ExcelCellUtils.excelIndexToCell(rowIndex, colIndex)));
                                }
                                if (skipEmptyCells.containsKey(rowIndex)
                                        && skipEmptyCells.get(rowIndex).contains(colIndex)) {
                                    continue;
                                }
                            }
                            Object apply = cellConverter.apply(cellValue);
                            cellConsumer.accept(currentNewObject, apply);
                            setData((R) currentNewObject);
                        } catch (Exception e) {
                            EXISTS_READ_FAIL.set(true);
                        }
                    }
                    /* SAX解析会跳过空白单元格，需显式检查缺失的注册单元格 */
                    for (Integer colIndex : cell2setter.get(rowIndex).keySet()) {
                        if (headMap.containsKey(colIndex)) {
                            continue;
                        }
                        templateConsistencyCheckCells.remove(ExcelCellUtils.excelIndexToCell(rowIndex, colIndex));
                        try {
                            if (mustExistCells.containsKey(rowIndex)
                                    && mustExistCells.get(rowIndex).contains(colIndex)) {
                                throw new CheckedExcelException(
                                        String.format(ExcelConstants.FIELD_LACK_MSG, ExcelCellUtils.excelIndexToCell(rowIndex, colIndex)));
                            }
                            if (skipEmptyCells.containsKey(rowIndex)
                                    && skipEmptyCells.get(rowIndex).contains(colIndex)) {
                                continue;
                            }
                            ExFunction cellConverter = cell2converts.get(rowIndex).get(colIndex);
                            BiConsumer cellConsumer = cell2setter.get(rowIndex).get(colIndex);
                            Object apply = cellConverter.apply("");
                            cellConsumer.accept(currentNewObject, apply);
                            setData((R) currentNewObject);
                        } catch (Exception e) {
                            EXISTS_READ_FAIL.set(true);
                        }
                    }
                }

                @Override
                public void invoke(Map<Integer, Object> data, AnalysisContext context) {
                    int rowIndex = context.readRowHolder().getRowIndex();
                    if (cell2setter.size() <= 0 || !cell2setter.containsKey(rowIndex)) {
                        return;
                    }
                    for (Map.Entry<Integer, Object> entry : data.entrySet()) {
                        int colIndex = entry.getKey();
                        Object value = entry.getValue();

                        if (!cell2setter.get(rowIndex).containsKey(colIndex)) {
                            continue;
                        }

                        /* 模板一致性检查：命中后移除 */
                        templateConsistencyCheckCells.remove(ExcelCellUtils.excelIndexToCell(rowIndex, colIndex));

                        try {
                            ExFunction cellConverter = cell2converts.get(rowIndex).get(colIndex);
                            BiConsumer cellConsumer = cell2setter.get(rowIndex).get(colIndex);
                            String cellValue = StringConverter.parse(value);

                            /* 必填/选填检查 */
                            if (cellValue == null || cellValue.isEmpty()) {
                                if (mustExistCells.containsKey(rowIndex)
                                        && mustExistCells.get(rowIndex).contains(colIndex)) {
                                    throw new CheckedExcelException(
                                            String.format(ExcelConstants.FIELD_LACK_MSG, ExcelCellUtils.excelIndexToCell(rowIndex, colIndex)));
                                }
                                if (skipEmptyCells.containsKey(rowIndex)
                                        && skipEmptyCells.get(rowIndex).contains(colIndex)) {
                                    continue;
                                }
                            }

                            Object apply = cellConverter.apply(cellValue);
                            cellConsumer.accept(currentNewObject, apply);
                            setData((R) currentNewObject);
                        } catch (Exception e) {
                            EXISTS_READ_FAIL.set(true);
                        }
                    }

                    /* SAX解析会跳过空白单元格，需显式检查缺失的注册单元格 */
                    for (Integer colIndex : cell2setter.get(rowIndex).keySet()) {
                        if (data.containsKey(colIndex)) {
                            continue;
                        }
                        templateConsistencyCheckCells.remove(ExcelCellUtils.excelIndexToCell(rowIndex, colIndex));
                        try {
                            if (mustExistCells.containsKey(rowIndex)
                                    && mustExistCells.get(rowIndex).contains(colIndex)) {
                                throw new CheckedExcelException(
                                        String.format(ExcelConstants.FIELD_LACK_MSG, ExcelCellUtils.excelIndexToCell(rowIndex, colIndex)));
                            }
                            if (skipEmptyCells.containsKey(rowIndex)
                                    && skipEmptyCells.get(rowIndex).contains(colIndex)) {
                                continue;
                            }
                            ExFunction cellConverter = cell2converts.get(rowIndex).get(colIndex);
                            BiConsumer cellConsumer = cell2setter.get(rowIndex).get(colIndex);
                            Object apply = cellConverter.apply("");
                            cellConsumer.accept(currentNewObject, apply);
                            setData((R) currentNewObject);
                        } catch (Exception e) {
                            EXISTS_READ_FAIL.set(true);
                        }
                    }
                }

                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    // 处理从未被触发过的行（整行空白或被SAX跳过）
                    for (String cellRef : new HashSet<>(templateConsistencyCheckCells)) {
                        int rowIndex = ExcelCellUtils.excelCellToRowIndex(cellRef);
                        int colIndex = ExcelCellUtils.excelCellToColumnIndex(cellRef);
                        templateConsistencyCheckCells.remove(cellRef);
                        try {
                            if (mustExistCells.containsKey(rowIndex)
                                    && mustExistCells.get(rowIndex).contains(colIndex)) {
                                throw new CheckedExcelException(
                                        String.format(ExcelConstants.FIELD_LACK_MSG, cellRef));
                            }
                            if (skipEmptyCells.containsKey(rowIndex)
                                    && skipEmptyCells.get(rowIndex).contains(colIndex)) {
                                continue;
                            }
                            ExFunction cellConverter = cell2converts.get(rowIndex).get(colIndex);
                            BiConsumer cellConsumer = cell2setter.get(rowIndex).get(colIndex);
                            Object apply = cellConverter.apply("");
                            cellConsumer.accept(currentNewObject, apply);
                            setData((R) currentNewObject);
                        } catch (Exception e) {
                            EXISTS_READ_FAIL.set(true);
                        }
                    }
                }
            }).sheet(sheetName).doRead();

            log.debug("[EasyExcelCellReader] Sheet读取完成: rId={}, sheetName={}",
                    sheetRid, sheetName);
        } catch (Exception e) {
            log.error("EasyExcel读取Excel文件异常", e);
        } finally {
            ExcelUtilsConfig.importParallelSemaphore.release();
        }
        return getData();
    }
}
