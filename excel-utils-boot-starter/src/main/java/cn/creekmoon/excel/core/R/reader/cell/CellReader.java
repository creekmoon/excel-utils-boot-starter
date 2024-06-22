package cn.creekmoon.excel.core.R.reader.cell;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.reader.Reader;
import cn.creekmoon.excel.util.exception.ExFunction;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Set;
import java.util.function.BiConsumer;


/**
 * This interface extends the `Reader` interface and provides methods to add cell converters for reading
 * and extracting data from specific cells in a spreadsheet. It allows converting values from string to
 * specified types and setting the converted values into corresponding fields or properties of an object.
 *
 * @param <R> the type of the object being read and populated with data
 */
public abstract class CellReader<R> extends Reader<R> {


    public Object currentNewObject;


    /* 必填项过滤  key=rowIndex  value=<colIndex> */
    public LinkedHashMap<Integer, Set<Integer>> mustExistCells = new LinkedHashMap<>(32);
    /* 选填项过滤  key=rowIndex  value=<colIndex> */
    public LinkedHashMap<Integer, Set<Integer>> skipEmptyCells = new LinkedHashMap<>(32);

    /* key=rowIndex  value=<colIndex,Consumer> 单元格转换器*/
    public LinkedHashMap<Integer, HashMap<Integer, ExFunction>> cell2converts = new LinkedHashMap(32);

    /* key=rowIndex  value=<colIndex,Consumer> 单元格消费者(通常是setter方法)*/
    public LinkedHashMap<Integer, HashMap<Integer, BiConsumer>> cell2setter = new LinkedHashMap(32);

    /*启用模板一致性检查 为了防止模板导入错误*/
    public boolean TEMPLATE_CONSISTENCY_CHECK_ENABLE = true;

    /*标志位, 模板一致性检查已经失败 */
    protected boolean TEMPLATE_CONSISTENCY_CHECK_HAS_FAILED = false;

    public CellReader(ExcelImport parent) {
        super(parent);
    }


    /**
     * 添加一个单元格转换器
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param convert       数值类型适配器, 例如 String --> Date
     * @param setter        Setter方法, 例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    abstract public <T> CellReader<R> addConvert(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param reader        Setter方法, 例如 setName(String name)
     * @return
     */
    abstract public CellReader<R> addConvert(String cellReference, BiConsumer<R, String> reader);

    /**
     * 添加一个单元格转换器
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param convert  数值类型适配器, 例如 String --> Date
     * @param setter   Setter方法, 例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    abstract public <T> CellReader<R> addConvert(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param setter   Setter方法, 例如 setName(String name)
     * @return
     */
    abstract public CellReader<R> addConvert(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    /**
     * 添加一个单元格转换器并跳过空值
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param setter   Setter方法, 例如 setName(String name)
     * @param <T>
     * @return
     */
    abstract public <T> CellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    /**
     * 添加一个单元格转换器并跳过空值
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param convert  数值类型适配器,例如 String --> Date
     * @param setter   Setter方法,例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    abstract public <T> CellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器并跳过空值
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param convert       数值类型适配器,例如 String --> Date
     * @param setter        Setter方法,例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    abstract public <T> CellReader<R> addConvertAndSkipEmpty(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器并跳过空值
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param setter        Setter方法,例如 setName(String name)
     * @return
     */
    abstract public CellReader<R> addConvertAndSkipEmpty(String cellReference, BiConsumer<R, String> setter);

    /**
     * 添加一个单元格转换器并要求存在值
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param convert  数值类型适配器,例如 String -> Date
     * @param setter   Setter方法,例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    abstract public <T> CellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器并要求存在值
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param setter   Setter方法,例如 setName(String name)
     * @return
     */
    abstract public CellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    /**
     * 添加一个单元格转换器并要求存在值
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param setter        Setter方法,例如 setName(String name)
     * @return
     */
    abstract public CellReader<R> addConvertAndMustExist(String cellReference, BiConsumer<R, String> setter);

}


