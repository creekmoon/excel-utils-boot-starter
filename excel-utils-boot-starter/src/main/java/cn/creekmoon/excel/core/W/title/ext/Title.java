package cn.creekmoon.excel.core.W.title.ext;

import cn.hutool.core.util.ArrayUtil;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.function.Function;

/**
 * 标题类  是个单向链表 会指向自己的父表头
 *
 * @param
 */
public class Title {


    public Title parentTitle;
    public String titleName;
    public Function valueFunction;
    /* 列坐标 开始于第几列  */
    public int startColIndex;
    /* 列坐标 结束于第几列 */
    public int endColIndex;
    /* 父标题分隔符*/
    public String PARENT_TITLE_SEPARATOR = "::";
    /* 携带的单元格样式 */
    public List<ConditionCellStyle> conditionCellStyleList = new ArrayList<>();


    private Title() {

    }


    public static Title of(String titleName, Function valueFunction, ConditionCellStyle... conditionCellStyles) {
        Title newTitle = new Title();
        newTitle.titleName = titleName;
        newTitle.valueFunction = valueFunction;
        if (conditionCellStyles != null && conditionCellStyles.length > 0) {
            newTitle.conditionCellStyleList.addAll(Arrays.asList(conditionCellStyles));
        }
        return newTitle;
    }


    /**
     * (如果可以)将Title转换成链表  并返回一个深度集合   每个Title指向一个parentTitle
     *
     * @param currentColIndex 当前标题的列位置
     * @return Map key=深度  value = title对象
     */
    public HashMap<Integer, Title> convert2ChainTitle(int currentColIndex) {
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
                    Title parentTitle = new Title();
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
