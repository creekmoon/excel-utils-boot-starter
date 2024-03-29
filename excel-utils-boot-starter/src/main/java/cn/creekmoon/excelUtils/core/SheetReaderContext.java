package cn.creekmoon.excelUtils.core;

import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

public class SheetReaderContext {
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

    public int sheetIndex;
    public Supplier newObjectSupplier;
    public Object currentNewObject;

    StringBuilder errorReport = new StringBuilder();
    /*单元格转换后置处理*/
    protected List<ExConsumer> cellConvertPostProcessors = new ArrayList<>();
    /* 必填项过滤  key=rowIndex  value=<colIndex> */
    protected LinkedHashMap<Integer, Set<Integer>> mustExistCells = new LinkedHashMap<>(32);
    /* 选填项过滤  key=rowIndex  value=<colIndex> */
    protected LinkedHashMap<Integer, Set<Integer>> skipEmptyCells = new LinkedHashMap<>(32);
    /* key=rowIndex  value=<colIndex,Consumer> 单元格转换器*/
    protected LinkedHashMap<Integer, HashMap<Integer, ExFunction>> cell2converts = new LinkedHashMap(32);

    /* key=rowIndex  value=<colIndex,Consumer> 单元格消费者*/
    protected LinkedHashMap<Integer, HashMap<Integer, BiConsumer>> cell2consumers = new LinkedHashMap(32);



    /* key=title  value=执行器 */
    protected LinkedHashMap<String, ExFunction> title2converts = new LinkedHashMap(32);
    /* key=title value=消费者(通常是setter方法)*/
    protected LinkedHashMap<String, BiConsumer> title2consumers = new LinkedHashMap(32);
    protected List<ExConsumer> convertPostProcessors = new ArrayList<>();
    /* key=title */
    protected Set<String> mustExistTitles = new HashSet<>(32);
    protected Set<String> skipEmptyTitles = new HashSet<>(32);
    /*启用空白行过滤*/
    protected boolean ENABLE_BLANK_ROW_FILTER = true;
    /*启用EXCEL标题模板一致性检查 为了防止模板导入错误*/
    protected boolean ENABLE_TITLE_CHECK = true;
    /*标志位,如果标题检查失败, 这个会置为true */
    protected boolean TITLE_CHECK_FAIL_FLAG = false;

    public SheetReaderContext(int sheetIndex, Supplier newObjectSupplier) {
        this.sheetIndex = sheetIndex;
        this.newObjectSupplier = newObjectSupplier;
    }
}
