package cn.creekmoon.excel.core.W.title;

import cn.creekmoon.excel.core.W.Writer;
import cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle;
import cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle;
import cn.creekmoon.excel.core.W.title.ext.Title;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.function.Function;
import java.util.function.Predicate;

public abstract class TitleWriter<R> extends Writer {

    /**
     * 表头集合
     */
    protected List<Title> titles = new ArrayList<>();


    /* 多级表头时会用到 全局标题深度  initTitle方法会给其赋值
     *
     *   |       title               |     深度=3    rowIndex=0
     *   |   titleA    |    titleB   |     深度=2    rowIndex=1
     *   |title1|title2|title3|title4|     深度=1    rowIndex=2
     * */
    protected Integer MAX_TITLE_DEPTH = null;

    /* 多级表头时会用到  深度和标题的映射关系*/
    protected HashMap<Integer, List<Title>> depth2Titles = new HashMap<>();

    /**
     * 标题行号 这里是0,意味着第一行是标题
     */
    protected int titleRowIndex = 0;

    /**
     * 首行数据行号
     */
    protected int firstRowIndex = titleRowIndex + 1;

    /**
     * 末行数据行号
     */
    protected int latestRowIndex = Integer.MAX_VALUE;


    /**
     * 添加标题
     */
    public TitleWriter<R> addTitle(String titleName, Function<R, Object> valueFunction) {
        return addTitle(titleName, valueFunction, (ConditionCellStyle) null);
    }

    /**
     * 添加标题
     */
    public TitleWriter<R> addTitle(String titleName, Function<R, Object> valueFunction, ExcelCellStyle cellStyle, Predicate<R> cellStyleCondition) {
        return addTitle(titleName, valueFunction, ConditionCellStyle.of(cellStyle, cellStyleCondition));
    }

    /**
     * 添加标题
     */
    public TitleWriter<R> addTitle(String titleName, Function<R, Object> valueFunction, ConditionCellStyle<R>... conditionStyle) {
        titles.add(Title.of(titleName, valueFunction, conditionStyle));
        return this;
    }


    public abstract int countTitles();

    abstract protected void doWrite(List<R> data);

    public TitleWriter<R> write(List<R> data) {
        preWrite();
        this.doWrite(data);
        return this;
    }

    /**
     * 重置写入器以支持在同一个sheet中写入不同类型的表格
     * 新的写入器会自动从当前位置的下一行开始（留1行空白）
     * 可以通过 range() 方法重新指定写入位置
     *
     * @param newDataClass 新表格的数据类型
     * @param <T> 新的数据类型
     * @return 新的 TitleWriter 实例
     */
    abstract public <T> TitleWriter<T> reset(Class<T> newDataClass);

    /**
     * 增加写入范围限制
     *
     * @param titleRowIndex    标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param firstDataRowIndex 首条数据所在的行数(下标按照从0开始)
     * @param lastDataRowIndex 最后一条数据所在的行数(下标按照从0开始)
     * @return
     */
    abstract public TitleWriter<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex);

    /**
     * 增加写入范围限制
     *
     * @param startRowIndex 标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param lastRowIndex  最后一条数据所在的行数(下标按照从0开始)
     * @return
     */
    abstract public TitleWriter<R> range(int startRowIndex, int lastRowIndex);

    /**
     * 增加写入范围限制
     *
     * @param startRowIndex 起始行下标(从0开始)
     * @return
     */
    abstract public TitleWriter<R> range(int startRowIndex);

}
