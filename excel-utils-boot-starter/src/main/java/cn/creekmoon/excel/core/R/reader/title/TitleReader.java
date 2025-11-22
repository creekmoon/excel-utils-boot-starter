package cn.creekmoon.excel.core.R.reader.title;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.reader.Reader;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;

import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Supplier;


/**
 * CellReader
 * 遍历所有行, 并按标题读取
 * 一个sheet页只会有多个结果对象, 每行是一个对象
 *
 * @param <R>
 */
public abstract class TitleReader<R> extends Reader<R> {

    /**
     * 标题行号 这里是0,意味着第一行是标题
     */
    public int titleRowIndex = 0;

    /**
     * 首行数据行号
     */
    public int firstRowIndex = titleRowIndex + 1;

    /**
     * 末行数据行号
     */
    public int latestRowIndex = Integer.MAX_VALUE;

    public HashMap<Integer, String> colIndex2Title = new HashMap<>();

    /* key=title  value=执行器 */
    public LinkedHashMap<String, ExFunction> title2converts = new LinkedHashMap(32);
    /* key=title value=消费者(通常是setter方法)*/
    public LinkedHashMap<String, BiConsumer> title2consumers = new LinkedHashMap(32);
    public List<ExConsumer> convertPostProcessors = new ArrayList<>();
    /* key=title */
    public Set<String> mustExistTitles = new HashSet<>(32);
    public Set<String> skipEmptyTitles = new HashSet<>(32);

    /*启用空白行过滤*/
    public boolean ENABLE_BLANK_ROW_FILTER = true;

    public TitleReader(ExcelImport parent) {
        super(parent);
    }


    abstract public Long getSheetRowCount();

    abstract public <T> TitleReader<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    abstract public TitleReader<R> addConvert(String title, BiConsumer<R, String> reader);

    abstract public <T> TitleReader<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter);

    abstract public <T> TitleReader<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    abstract public TitleReader<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter);

    abstract public <T> TitleReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    abstract public <T> TitleReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor);

    abstract public TitleReaderResult<R> read(ExConsumer<R> dataConsumer);

    abstract public TitleReaderResult<R> read();

    abstract public TitleReader<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex);

    abstract public TitleReader<R> range(int startRowIndex, int lastRowIndex);

    abstract public TitleReader<R> range(int startRowIndex);

    abstract public Integer getSheetIndex();

    /**
     * 重置读取器以支持在同一个sheet中读取不同类型的表格
     * 新的读取器会清空所有转换规则和范围设置
     * 需要重新调用 addConvert() 和 range() 配置
     *
     * @param newObjectSupplier 新表格的对象创建器
     * @param <T> 新的数据类型
     * @return 新的 TitleReader 实例
     */
    abstract public <T> TitleReader<T> reset(Supplier<T> newObjectSupplier);



}
