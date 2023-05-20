package cn.creekmoon.excelUtils.core;

import cn.hutool.core.util.ArrayUtil;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.function.Function;

public class SheetWriterContext {

    protected String sheetName;

    /**
     * 表头集合
     */
    protected List<Title> titles = new ArrayList<>();
    /**
     * 表头集合
     */
    protected HashMap<Integer, List<SheetWriter.ConditionStyle>> colIndex2Styles = new HashMap<>();

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
     * 当前写入行,切换sheet页时需要还原这个上下文数据
     */
    protected int currentRow = 0;


    /**
     * 标题类  是个单向链表 会指向自己的父表头
     *
     * @param <R>
     */
    public static class Title<R> {
        public Title parentTitle;
        public String titleName;
        public Function<R, Object> valueFunction;
        /* 父标题分隔符*/
        public static String PARENT_TITLE_SEPARATOR = "::";
        /* 列坐标 开始于第几列  */
        public int startColIndex;
        /* 列坐标 结束于第几列 */
        public int endColIndex;

        public Title(String titleName, Function<R, Object> valueFunction) {
            this.titleName = titleName;
            this.valueFunction = valueFunction;
        }

        private Title() {

        }

        public static <R> Title<R> of(String titleName, Function<R, Object> valueFunction) {
            Title<R> newTitle = new Title<>();
            newTitle.titleName = titleName;
            newTitle.valueFunction = valueFunction;
            return newTitle;
        }

        /**
         * (如果可以)将Title转换成链表  并返回一个深度集合   每个Title指向一个parentTitle
         *
         * @param currentColIndex 当前标题的列位置
         * @return Map key=深度  value = title对象
         */
        protected HashMap<Integer, Title> convert2ChainTitle(int currentColIndex) {
            this.startColIndex = currentColIndex;
            this.endColIndex = currentColIndex;

            HashMap<Integer, Title> depth2Title = new HashMap<>();
            depth2Title.put(1, this);
            if (titleName.contains(PARENT_TITLE_SEPARATOR)) {
                //倒序  源: titleParentParent::titleParent::title  转换后 [title,titleParent,titleParentParent]
                String[] split = ArrayUtil.reverse(titleName.split(PARENT_TITLE_SEPARATOR));
                this.titleName = split[0];

                Title currentTitle = this;
                for (int i = 0; i < split.length; i++) {
                    int depth = i + 1;
                    if (depth != split.length) {
                        /* 生成父Title */
                        Title parentTitle = new Title<>();
                        parentTitle.titleName = split[depth];
                        parentTitle.startColIndex = currentColIndex;
                        parentTitle.endColIndex = currentColIndex;
                        currentTitle.parentTitle = parentTitle;
                    }
                    depth2Title.put(depth, currentTitle);
                    currentTitle = currentTitle.parentTitle;
                }
            }
            return depth2Title;
        }
    }

    public SheetWriterContext(String sheetName) {
        this.sheetName = sheetName;
    }
}
