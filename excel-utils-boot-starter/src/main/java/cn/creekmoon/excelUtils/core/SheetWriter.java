package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.core.SheetWriterContext.Title;
import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.style.StyleUtil;
import jakarta.servlet.http.HttpServletResponse;
import lombok.AllArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.io.IOException;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.Collectors;

@Slf4j
public class SheetWriter<R> {

    public SheetWriterContext sheetWriterContext;

    protected ExcelExport parent;


    /**
     * 添加标题
     */
    public SheetWriter<R> addTitle(String titleName, Function<R, Object> valueFunction) {
        sheetWriterContext.titles.add(Title.of(titleName, valueFunction));
        return this;
    }

    /**
     * 获取当前的表头数量
     */
    public int countTitles() {
        return sheetWriterContext.titles.size();
    }

    /**
     * 为当前列设置一个样式
     *
     * @param condition        样式触发的条件
     * @param styleInitializer 样式初始化器
     * @return
     */
    public SheetWriter<R> setDataStyle(Predicate<R> condition, Consumer<XSSFCellStyle> styleInitializer) {
        return setDataStyle(sheetWriterContext.titles.size() - 1, condition, styleInitializer);
    }

    /**
     * 添加excel样式映射,在写入的时候会读取这个映射
     *
     * @param colIndex         对应的列号
     * @param condition        样式的触发条件
     * @param styleInitializer 样式的初始化内容
     * @return
     */
    public SheetWriter<R> setDataStyle(int colIndex, Predicate<R> condition, Consumer<XSSFCellStyle> styleInitializer) {
        /*初始化样式*/
//        CellStyle newCellStyle = parent.getBigExcelWriter().createCellStyle();
        XSSFCellStyle newCellStyle = (XSSFCellStyle) StyleUtil.createDefaultCellStyle(parent.getBigExcelWriter().getWorkbook());
        styleInitializer.accept(newCellStyle);
        ConditionStyle conditionStyle = new ConditionStyle(condition, newCellStyle);

        /*保存映射结果*/
        if (!sheetWriterContext.colIndex2Styles.containsKey(colIndex)) {
            sheetWriterContext.colIndex2Styles.put(colIndex, new ArrayList<>());
        }
        sheetWriterContext.colIndex2Styles.get(colIndex).add(conditionStyle);
        return this;
    }


    /**
     * 为当前列的数据设置一个样式
     *
     * @param styleInitializer 样式初始化器
     * @return
     */
    public SheetWriter<R> setDataStyle(Consumer<XSSFCellStyle> styleInitializer) {
        return this.setDataStyle(x -> true, styleInitializer);
    }

    /**
     * 写入对象
     *
     * @param data
     * @return
     */
    public SheetWriter<R> write(List<R> data) {
        return write(data, ExcelExport.WriteStrategy.CONTINUE_ON_ERROR);
    }


    /**
     * 以map形式写入
     *
     * @param data key=标题 value=行内容
     * @return
     */
    protected SheetWriter<R> writeByMap(List<Map<String, Object>> data) {
        preWritingContextAdjust();
        parent.getBigExcelWriter().write(data);
        postWritingContextAdjust();
        return this;
    }

    /**
     * 以对象形式写入
     *
     * @param vos           数据集
     * @param writeStrategy 写入策略
     * @return
     */
    private SheetWriter<R> write(List<R> vos, ExcelExport.WriteStrategy writeStrategy) {
        preWritingContextAdjust();

        this.initTitles();
        List<List<Object>> rows =
                vos.stream()
                        .map(
                                vo -> {
                                    List<Object> row = new LinkedList<>();
                                    for (Title<R> title : sheetWriterContext.titles) {
                                        //当前行中的某个属性值
                                        Object data = null;
                                        try {
                                            data = title.valueFunction.apply(vo);
                                        } catch (Exception exception) {
                                            if (writeStrategy == ExcelExport.WriteStrategy.CONTINUE_ON_ERROR) {
                                                // nothing to do
                                                if (parent.debugger) {
                                                    log.info("[Excel构建]生成Excel获取数据值时发生错误!已经忽略错误并设置为NULL值!", exception);
                                                }
                                            }
                                            if (writeStrategy == ExcelExport.WriteStrategy.STOP_ON_ERROR) {
                                                String taskId = parent.stopWrite();
                                                ExcelExport.cleanTempFileDelay(taskId);
                                                log.error("[Excel构建]生成Excel获取数据值时发生错误!", exception);
                                                throw new RuntimeException("生成Excel获取数据值时发生错误!");
                                            }
                                        }
                                        row.add(data);
                                    }
                                    return row;
                                })
                        .collect(Collectors.toList());

        /*一口气写入, 让内部组件自己自动分批写到磁盘  区别是这样就不能设置style了*/
        //parent.getBigExcelWriter().write(rows);

        /*将传入过来的rows按每批100次进行写入到硬盘 此时可以设置style, 底层不会报"已经写入磁盘无法编辑"的异常*/
        List<List<R>> subVos = ListUtil.partition(vos, BigExcelWriter.DEFAULT_WINDOW_SIZE);
        List<List<List<Object>>> subRows = ListUtil.partition(rows, BigExcelWriter.DEFAULT_WINDOW_SIZE);
        for (int i = 0; i < subRows.size(); i++) {
            int startRowIndex = parent.getBigExcelWriter().getCurrentRow();
            parent.getBigExcelWriter().write(subRows.get(i));
            int endRowIndex = parent.getBigExcelWriter().getCurrentRow();
            setDataStyle(subVos.get(i), startRowIndex, endRowIndex);
        }

        postWritingContextAdjust();
        return this;
    }


    /**
     * 给excel设置样式
     *
     * @param vos
     * @param startRowIndex
     * @param endRowIndex
     */
    private void setDataStyle(List<R> vos, int startRowIndex, int endRowIndex) {
        for (int i = 0; i < vos.size(); i++) {
            R vo = vos.get(i);
            for (int colIndex = 0; colIndex < sheetWriterContext.titles.size(); colIndex++) {
                List<ConditionStyle> styleList = sheetWriterContext.colIndex2Styles.getOrDefault(colIndex, Collections.EMPTY_LIST);
                for (ConditionStyle conditionStyle : styleList) {
                    if (conditionStyle.condition.test(vo)) {
                        /*写回样式*/
                        parent.getBigExcelWriter().setStyle(conditionStyle.style, colIndex, startRowIndex + i);
                    }
                }
            }
        }
    }


    /**
     * 初始化标题
     */
    private void initTitles() {

        /*如果已经初始化完毕 则不进行初始化*/
        if (isTitleInitialized()) {
            return;
        }

        sheetWriterContext.MAX_TITLE_DEPTH = sheetWriterContext.titles.stream()
                .map(x -> StrUtil.count(x.titleName, Title.PARENT_TITLE_SEPARATOR) + 1)
                .max(Comparator.naturalOrder())
                .orElse(1);
        if (parent.debugger) {
            System.out.println("[Excel构建] 表头深度获取成功! 表头最大深度为" + sheetWriterContext.MAX_TITLE_DEPTH);
        }

        /*多级表头初始化*/
        for (int i = 0; i < sheetWriterContext.titles.size(); i++) {
            Title oneTitle = sheetWriterContext.titles.get(i);
            changeTitleWithMaxlength(oneTitle);
            HashMap<Integer, Title> map = oneTitle.convert2ChainTitle(i);
            map.forEach((depth, title) -> {
                if (!sheetWriterContext.depth2Titles.containsKey(depth)) {
                    sheetWriterContext.depth2Titles.put(depth, new ArrayList<>());
                }
                sheetWriterContext.depth2Titles.get(depth).add(title);
            });
        }

        /*横向合并相同名称的标题 PS:不会对最后一行进行横向合并 意味着允许最后一行出现相同的名称*/
        for (int i = 1; i <= sheetWriterContext.MAX_TITLE_DEPTH; i++) {
            Integer rowsIndex = getRowsIndexByDepth(i);
            final int finalI = i;
            Map<Object, List<Title>> collect = sheetWriterContext.depth2Titles
                    .get(i)
                    .stream()
                    .collect(Collectors.groupingBy(title -> {
                        /*如果深度为1则不分组*/
                        if (finalI == 1) {
                            return title;
                        }
                        /*如果深度大于1则分组 同一组的就进行横向合并*/
                        Title x = title;
                        StringBuilder groupName = new StringBuilder(x.titleName);
                        while (x.parentTitle != null) {
                            groupName.insert(0, x.parentTitle.titleName + Title.PARENT_TITLE_SEPARATOR);
                            x = x.parentTitle;
                        }
                        return groupName.toString();
                    }, Collectors.toList()));
            collect
                    .values()
                    .forEach(list -> {
                        int startColIndex = list.stream().map(x -> x.startColIndex).min(Comparator.naturalOrder()).orElse(1);
                        int endColIndex = list.stream().map(x -> x.endColIndex).max(Comparator.naturalOrder()).orElse(1);
                        if (startColIndex == endColIndex) {
                            if (parent.debugger) {
                                System.out.println("[Excel构建] 插入表头[" + list.get(0).titleName + "]");
                            }
                            this.parent.getBigExcelWriter().getOrCreateCell(startColIndex, rowsIndex).setCellValue(list.get(0).titleName);
                            this.parent.getBigExcelWriter().getOrCreateCell(startColIndex, rowsIndex).setCellStyle(this.parent.getBigExcelWriter().getHeadCellStyle());
                            return;
                        }
                        if (parent.debugger) {
                            System.out.println("[Excel构建] 插入表头并横向合并[" + list.get(0).titleName + "]预计合并格数[" + (endColIndex - startColIndex) + "]");
                        }
                        this.parent.getBigExcelWriter().merge(rowsIndex, rowsIndex, startColIndex, endColIndex, list.get(0).titleName, true);
                        if (parent.debugger) {
                            System.out.println("[Excel构建] 横向合并表头[" + list.get(0).titleName + "]完毕");
                        }
                    });
            /* hutool excel 写入下移一行*/
            this.parent.getBigExcelWriter().setCurrentRow(this.parent.getBigExcelWriter().getCurrentRow() + 1);
        }

        /*纵向合并title*/
        for (int colIndex = 0; colIndex < sheetWriterContext.titles.size(); colIndex++) {
            Title<R> title = sheetWriterContext.titles.get(colIndex);
            int sameCount = 0; //重复的数量
            for (Integer depth = 1; depth <= sheetWriterContext.MAX_TITLE_DEPTH; depth++) {
                if (title.parentTitle != null && title.titleName.equals(title.parentTitle.titleName)) {
                    /*发现重复的单元格,不会马上合并 因为可能有更多重复的*/
                    sameCount++;
                } else if (sameCount != 0) {
                    if (parent.debugger) {
                        System.out.println("[Excel构建] 表头纵向合并" + title.titleName + "  预计合并格数" + sameCount);
                    }
                    /*合并单元格*/
                    this.parent.getBigExcelWriter().merge(getRowsIndexByDepth(depth), getRowsIndexByDepth(depth) + sameCount, colIndex, colIndex, title.titleName, true);
                    sameCount = 0;
                }
                title = title.parentTitle;
            }
        }

        this.setColumnWidthDefault();
    }


    /**
     * 是否已经初始化好表头
     *
     * @return
     */
    private boolean isTitleInitialized() {
        return this.sheetWriterContext.MAX_TITLE_DEPTH != null;
    }

    /**
     * 重置标题 当需要再次使用标题时
     */
    private void restTitles() {
        this.sheetWriterContext.MAX_TITLE_DEPTH = null;
        this.sheetWriterContext.titles.clear();
        this.sheetWriterContext.depth2Titles.clear();
    }


    /**
     * 默认设置列宽 : 简单粗暴将前500列都设置宽度20
     */
    public SheetWriter<R> setColumnWidthDefault() {
        this.setColumnWidth(500, 20);
        return this;
    }

    /**
     * 自动设置列宽
     */
    public void setColumnWidthAuto() {
        try {
            parent.getBigExcelWriter().autoSizeColumnAll();
        } catch (Exception ignored) {
        }
    }

    /**
     * 默认设置列宽 : 简单粗暴将前500列都设置宽度
     */
    public void setColumnWidth(int cols, int width) {
        for (int i = 0; i < cols; i++) {
            try {
                parent.getBigExcelWriter().setColumnWidth(i, width);
            } catch (Exception ignored) {
            }
        }
    }

    /**
     * 获取行坐标   如果最大深度为4  当前深度为1  则处于第4行
     *
     * @param depth 深度
     * @return
     */
    private Integer getRowsIndexByDepth(int depth) {
        return sheetWriterContext.MAX_TITLE_DEPTH - depth;
    }


    /**
     * 自动复制最后一级的表头  以补足最大表头深度  这样可以使深度统一
     * <p>
     * 例如:   A::B::C
     * D::F::G::H
     * 这样表头会参差不齐, 对A::B::C 补全为 A::B::C::C 统一成为了4级深度
     *
     * @param title
     * @return 原对象
     */
    private Title changeTitleWithMaxlength(Title title) {
        int currentMaxDepth = StrUtil.count(title.titleName, Title.PARENT_TITLE_SEPARATOR) + 1;
        if (sheetWriterContext.MAX_TITLE_DEPTH != currentMaxDepth) {
            String[] split = title.titleName.split(Title.PARENT_TITLE_SEPARATOR);
            String lastTitleName = split[split.length - 1];
            for (int y = 0; y < sheetWriterContext.MAX_TITLE_DEPTH - currentMaxDepth; y++) {
                title.titleName = title.titleName + Title.PARENT_TITLE_SEPARATOR + lastTitleName;
            }
        }
        return title;
    }


    protected SheetWriter(ExcelExport parent, SheetWriterContext sheetWriterContext) {
        this.parent = parent;
        this.sheetWriterContext = sheetWriterContext;
    }


    public String stopWrite() throws IOException {
        return parent.stopWrite();

    }


    /**
     * 切换到新的标签页
     */
    public <T> SheetWriter<T> switchSheet(String sheetName, Class<T> newDataClass) {
        return parent.switchSheet(sheetName, newDataClass);
    }

    /**
     * 写前上下文调整
     */
    public void preWritingContextAdjust() {
        if (parent.getBigExcelWriter().getSheet().getSheetName().equals(sheetWriterContext.sheetName)) {
            return;
        }
        parent.getBigExcelWriter().setSheet(sheetWriterContext.sheetName);
        parent.getBigExcelWriter().setCurrentRow(sheetWriterContext.currentRow);
    }

    /**
     * 写后上下文调整
     */
    public void postWritingContextAdjust() {
        sheetWriterContext.currentRow = parent.getBigExcelWriter().getCurrentRow();
    }

    /**
     * 切换到新的标签页
     */
    public <T> SheetWriter<T> switchSheet(Class<T> newDataClass) {
        return parent.switchSheet(newDataClass);
    }

    /**
     * 响应并清除文件
     *
     * @param response
     * @throws IOException
     */
    public void response(HttpServletResponse response) throws IOException {
        String taskId = parent.stopWrite();
        ExcelExport.response(taskId, parent.excelName, response);
    }

    /*条件样式*/
    @AllArgsConstructor
    public class ConditionStyle {
        Predicate<R> condition;
        CellStyle style;
    }
}
