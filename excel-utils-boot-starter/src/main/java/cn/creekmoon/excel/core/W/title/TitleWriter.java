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
        onWrite();
        this.doWrite(data);
        return this;
    }


}
