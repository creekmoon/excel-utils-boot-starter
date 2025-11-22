package cn.creekmoon.excel.core.R.reader.title;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.reader.Reader;
import cn.creekmoon.excel.util.exception.ExBiConsumer;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;
import cn.creekmoon.excel.util.exception.ExTriConsumer;
import lombok.SneakyThrows;

import java.util.*;
import java.util.concurrent.atomic.AtomicReference;
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

    /*启用模板一致性检查 为了防止模板导入错误*/
    public boolean ENABLE_TEMPLATE_CONSISTENCY_REVIEW = true;


    /*行结果集合*/
    public LinkedHashMap<Integer, String> rowIndex2msg = new LinkedHashMap<>();

    /*存在读取失败的数据*/
    public AtomicReference<Boolean> EXISTS_READ_FAIL = new AtomicReference<>(false);

    public TitleReader(ExcelImport parent) {
        super(parent);
    }


    /**
     * 拿到所有的数据, 并标记导入成功.
     * @return
     */
    public List<R> getAll() {
        ArrayList<R> objects = new ArrayList<>(128);
        read((ExConsumer<R>) objects::add);
        return objects;
    }

    abstract public <T> TitleReader<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    abstract public TitleReader<R> addConvert(String title, BiConsumer<R, String> reader);

    abstract public <T> TitleReader<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter);

    abstract public <T> TitleReader<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    abstract public TitleReader<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter);

    abstract public <T> TitleReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    abstract public <T> TitleReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor);


    /**
     * 消费单条数据
     * 1.消费过程中没有抛出异常, 则认为导入成功
     * 2.消费过程中抛出异常, 终止导入
     * @param dataConsumer 数据消费器
     * @return
     */
    abstract public TitleReader<R> read(ExConsumer<R> dataConsumer);


    /**
     * 消费单条数据
     * 1.单条消费过程中没有抛出异常, 则认为导入成功
     * 2.单条消费过程中抛出异常, 则后续终止导入
     * @param dataConsumer 第一个入参是数据下标索引 第二个参数是数据本身
     * @return
     */
    public abstract TitleReader<R> read(ExBiConsumer<Integer, R> dataConsumer);

    abstract public TitleReader<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex);

    abstract public TitleReader<R> range(int startRowIndex, int lastRowIndex);

    abstract public TitleReader<R> range(int startRowIndex);

    /**
     * 重置读取器以支持在同一个sheet中读取不同类型的表格
     * 新的读取器会清空所有转换规则和范围设置
     * 需要重新调用 addConvert() 和 range() 配置
     *
     * @param newObjectSupplier 新表格的对象创建器
     * @param <T>               新的数据类型
     * @return 新的 TitleReader 实例
     */
    abstract public <T> TitleReader<T> reset(Supplier<T> newObjectSupplier);


}
