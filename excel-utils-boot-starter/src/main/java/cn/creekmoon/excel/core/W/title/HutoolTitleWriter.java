package cn.creekmoon.excel.core.W.title;

import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle;
import cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle;
import cn.creekmoon.excel.core.W.title.ext.Title;
import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.style.StyleUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.*;
import java.util.stream.Collectors;

@Slf4j
public class HutoolTitleWriter<R> extends TitleWriter<R> {


    /*写入器 hutoolTitleWriter专用*/
    public BigExcelWriter bigExcelWriter;
    protected ExcelExport parent;
    /**
     * 当前写入行,切换sheet页时需要还原这个上下文数据
     */
    protected int currentRow = 0;

    private String sheetInitKey(int sheetIndex) {
        return "SHEET_INIT_" + sheetIndex;
    }

    public HutoolTitleWriter(ExcelExport parent, Integer sheetIndex) {
        this.parent = parent;
        this.sheetIndex = sheetIndex;
        this.sheetName = "Sheet" + sheetIndex;
    }

    public HutoolTitleWriter(ExcelExport parent, Integer sheetIndex, String sheetName) {
        this.parent = parent;
        this.sheetIndex = sheetIndex;
        this.sheetName = sheetName;
    }

    /**
     * 重置写入器以支持在同一个sheet中写入不同类型的表格
     * 新的写入器会自动从当前位置的下一行开始
     *
     * @param newDataClass   新表格的数据类型
     * @param <T>            新的数据类型
     * @return 新的 TitleWriter 实例
     */
    @Override
    public <T> HutoolTitleWriter<T> reset(Class<T> newDataClass) {
        return reset(newDataClass, 0);
    }

    /**
     * 重置写入器以支持在同一个sheet中写入不同类型的表格
     * 新的写入器会自动从当前位置的下一行开始
     * 可以通过rowIndexOffset重新指定写入位置
     *
     * @param newDataClass   新表格的数据类型
     * @param rowIndexOffset 行偏移量 默认为0 (如果想隔一行再写入则设置1)
     * @param <T>            新的数据类型
     * @return 新的 TitleWriter 实例
     */
    @Override
    public <T> HutoolTitleWriter<T> reset(Class<T> newDataClass, int rowIndexOffset) {
        // 创建新的 writer 实例，但共享关键资源
        HutoolTitleWriter<T> newWriter = new HutoolTitleWriter<>(this.parent, this.sheetIndex, this.sheetName);

        // 共享 Excel 写入器（关键：确保写入同一个 Excel 文件和 sheet）
        newWriter.bigExcelWriter = this.bigExcelWriter;

        // 自动续接位置：从当前写入行的下一行开始
        newWriter.titleRowIndex = this.currentRow + rowIndexOffset;
        newWriter.firstRowIndex = null;  // 设置为 null，由 initTitles 根据表头深度自动推断
        newWriter.latestRowIndex = null; // 设置为 null，由 initTitles 设置为默认值

        // 继承当前行位置
        newWriter.currentRow = this.currentRow;

        return newWriter;
    }

    /**
     * 获取当前的表头数量
     */
    @Override
    public int countTitles() {
        return titles.size();
    }

    /**
     * 默认设置列宽 : 简单粗暴将前500列都设置宽度20
     */
    public HutoolTitleWriter<R> setColumnWidthDefault() {
        this.setColumnWidth(500, 20);
        return this;
    }

    /**
     * 自动设置列宽
     */
    public void setColumnWidthAuto() {
        try {
            getBigExcelWriter().autoSizeColumnAll();
        } catch (Exception ignored) {
        }
    }

    /**
     * 默认设置列宽 : 简单粗暴将前500列都设置宽度
     */
    public void setColumnWidth(int cols, int width) {
        for (int i = 0; i < cols; i++) {
            try {
                getBigExcelWriter().setColumnWidth(i, width);
            } catch (Exception ignored) {
            }
        }
    }


//    /**
//     * 应用预设的条件样式
//     *
//     * @param vos
//     * @param startRowIndex
//     * @param endRowIndex
//     */
//    private void applyConditionStyle(List<R> vos, int startRowIndex, int endRowIndex) {
//        for (int i = 0; i < vos.size(); i++) {
//            R vo = vos.get(i);
//            for (int colIndex = 0; colIndex < titles.size(); colIndex++) {
//                List<ConditionCellStyle> styleList = colIndex2Styles.getOrDefault(colIndex, Collections.EMPTY_LIST);
//                for (ConditionCellStyle conditionStyle : styleList) {
//                    if (conditionStyle.condition.test(vo)) {
//                        /*写回样式*/
//                        getBigExcelWriter().setStyle(parent.cellStyle2RunningTimeStyleObject.get(conditionStyle), colIndex, startRowIndex + i);
//                    }
//                }
//            }
//        }
//    }


    /**
     * 添加excel样式映射,在写入的时候会读取这个映射
     *
     * @param colIndex         对应的列号
     * @param condition        样式的触发条件
     * @param styleInitializer 样式的初始化内容
     * @return
     */
//    public HutoolTitleWriter<R> setDataStyle(int colIndex, Predicate<R> condition, Consumer<CellStyle> styleInitializer) {
//        /*初始化样式*/
////        CellStyle newCellStyle = getBigExcelWriter().createCellStyle();
//        CellStyle newCellStyle = StyleUtil.createDefaultCellStyle(getWorkbook());
//        styleInitializer.accept(newCellStyle);
//        ConditionCellStyle conditionStyle = new ConditionCellStyle(condition, newCellStyle);
//
//        /*保存映射结果*/
//        if (!colIndex2Styles.containsKey(colIndex)) {
//            colIndex2Styles.put(colIndex, new ArrayList<>());
//        }
//        colIndex2Styles.get(colIndex).add(conditionStyle);
//        return this;
//    }

    /**
     * 为当前列设置一个样式
     *
     * @param condition        样式触发的条件
     * @param styleInitializer 样式初始化器
     * @return
     */
//    public TitleWriter<R> setDataStyle(Predicate<R> condition, Consumer<CellStyle> styleInitializer) {
//        return setDataStyle(titles.size() - 1, condition, styleInitializer);
//    }

    /**
     * 增加写入范围限制
     *
     * @param titleRowIndex     标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param firstDataRowIndex 首条数据所在的行数(下标按照从0开始)
     * @param lastDataRowIndex  最后一条数据所在的行数(下标按照从0开始)
     * @return
     */
    @Override
    public HutoolTitleWriter<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex) {
        this.titleRowIndex = titleRowIndex;
        this.firstRowIndex = firstDataRowIndex;
        this.latestRowIndex = lastDataRowIndex;
        return this;
    }

    /**
     * 写入范围限制
     *
     * @param startRowIndex 标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param lastRowIndex  最后一条数据所在的行数(下标按照从0开始)
     * @return
     */
    @Override
    public HutoolTitleWriter<R> range(int startRowIndex, int lastRowIndex) {
        this.titleRowIndex = startRowIndex;
        this.firstRowIndex = null;  // 延迟到 initTitles 时根据表头深度自动推断
        this.latestRowIndex = lastRowIndex;
        return this;
    }

    /**
     * 写入范围限制
     *
     * @param startRowIndex 起始行下标(从0开始)
     * @return
     */
    @Override
    public HutoolTitleWriter<R> range(int startRowIndex) {
        this.titleRowIndex = startRowIndex;
        this.firstRowIndex = null;  // 延迟到 initTitles 时根据表头深度自动推断
        this.latestRowIndex = null; // 延迟到 initTitles 时设置为默认值
        return this;
    }

    /**
     * 内部操作类,但是暴露出来了,希望最好不要用这个方法
     *
     * @return
     */
    public BigExcelWriter getBigExcelWriter() {
        //如果上一个sheet页已经存在, 复用上一个sheet页的上下文
        if (bigExcelWriter == null && parent.sheetIndex2SheetWriter.containsKey(sheetIndex - 1)) {
            bigExcelWriter = ((HutoolTitleWriter) parent.sheetIndex2SheetWriter.get(sheetIndex - 1)).bigExcelWriter;
        }
        if (bigExcelWriter == null) {
            bigExcelWriter = ExcelUtil.getBigWriter(parent.getResultFilePath());
        }
        return bigExcelWriter;
    }

    @Override
    public Workbook getWorkbook() {
        return getBigExcelWriter().getWorkbook();
    }

    /**
     * 将对象转化为List<Object>形式, 每个下标对应一列
     *
     * @param dataObject
     * @return
     */
    private List<Object> changeToCellValues(R dataObject) {
        List<Object> row = new LinkedList<>();
        for (Title title : titles) {
            //当前行中的某个属性值
            Object data = null;
            try {
                data = title.valueFunction.apply(dataObject);
            } catch (Exception exception) {
                // nothing to do
                if (parent.debugger) {
                    log.info("[Excel构建]生成Excel获取数据值时发生错误!已经忽略错误并设置为NULL值!", exception);
                }
            }
            row.add(data);
        }
        return row;
    }

    /**
     * 初始化标题
     */
    private void initTitles() {

        /*如果已经初始化完毕 则不进行初始化*/
        if (isTitleInitialized()) {
            return;
        }

        /*设置默认值（如果用户未指定）*/
        if (titleRowIndex == null) {
            titleRowIndex = 0;
        }
        if (latestRowIndex == null) {
            latestRowIndex = Integer.MAX_VALUE;
        }

        MAX_TITLE_DEPTH = titles.stream()
                .map(x -> StrUtil.count(x.titleName, x.PARENT_TITLE_SEPARATOR) + 1)
                .max(Comparator.naturalOrder())
                .orElse(1);
        if (parent.debugger) {
            System.out.println("[Excel构建] 表头深度获取成功! 表头最大深度为" + MAX_TITLE_DEPTH);
        }

        /*🔥 核心推断逻辑：如果用户没有明确指定 firstRowIndex，则根据表头深度自动推断*/
        if (firstRowIndex == null) {
            firstRowIndex = titleRowIndex + MAX_TITLE_DEPTH;
            if (parent.debugger) {
                System.out.println("[Excel构建] 自动推断数据起始行: titleRowIndex=" + titleRowIndex
                        + ", MAX_TITLE_DEPTH=" + MAX_TITLE_DEPTH + ", firstRowIndex=" + firstRowIndex);
            }
        }

        /*多级表头初始化*/
        for (int i = 0; i < titles.size(); i++) {
            Title oneTitle = titles.get(i);
            changeTitleWithMaxlength(oneTitle);
            HashMap<Integer, Title> map = oneTitle.convert2ChainTitle(i);
            map.forEach((depth, title) -> {
                if (!depth2Titles.containsKey(depth)) {
                    depth2Titles.put(depth, new ArrayList<>());
                }
                depth2Titles.get(depth).add(title);
            });
        }

        /*横向合并相同名称的标题 PS:不会对最后一行进行横向合并 意味着允许最后一行出现相同的名称*/
        for (int i = 1; i <= MAX_TITLE_DEPTH; i++) {
            Integer rowsIndex = getRowsIndexByDepth(i);
            final int finalI = i;
            Map<Object, List<Title>> collect = depth2Titles
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
                            groupName.insert(0, x.parentTitle.titleName + x.PARENT_TITLE_SEPARATOR);
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
                            this.getBigExcelWriter().getOrCreateCell(startColIndex, rowsIndex).setCellValue(list.get(0).titleName);
                            this.getBigExcelWriter().getOrCreateCell(startColIndex, rowsIndex).setCellStyle(this.getBigExcelWriter().getHeadCellStyle());
                            return;
                        }
                        if (parent.debugger) {
                            System.out.println("[Excel构建] 插入表头并横向合并[" + list.get(0).titleName + "]预计合并格数[" + (endColIndex - startColIndex) + "]");
                        }
                        this.getBigExcelWriter().merge(rowsIndex, rowsIndex, startColIndex, endColIndex, list.get(0).titleName, true);
                        if (parent.debugger) {
                            System.out.println("[Excel构建] 横向合并表头[" + list.get(0).titleName + "]完毕");
                        }
                    });
            /* hutool excel 写入下移一行*/
            this.getBigExcelWriter().setCurrentRow(this.getBigExcelWriter().getCurrentRow() + 1);
        }

        /*纵向合并title*/
        for (int colIndex = 0; colIndex < titles.size(); colIndex++) {
            Title title = titles.get(colIndex);
            int sameCount = 0; //重复的数量
            for (Integer depth = 1; depth <= MAX_TITLE_DEPTH; depth++) {
                if (title.parentTitle != null && title.titleName.equals(title.parentTitle.titleName)) {
                    /*发现重复的单元格,不会马上合并 因为可能有更多重复的*/
                    sameCount++;
                } else if (sameCount != 0) {
                    if (parent.debugger) {
                        System.out.println("[Excel构建] 表头纵向合并" + title.titleName + "  预计合并格数" + sameCount);
                    }
                    /*合并单元格*/
                    this.getBigExcelWriter().merge(getRowsIndexByDepth(depth), getRowsIndexByDepth(depth) + sameCount, colIndex, colIndex, title.titleName, true);
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
        return this.MAX_TITLE_DEPTH != null;
    }

    /**
     * 重置标题 当需要再次使用标题时
     */
    private void restTitles() {
        this.MAX_TITLE_DEPTH = null;
        this.titles.clear();
        this.depth2Titles.clear();
    }

    /**
     * 获取行坐标   如果最大深度为4  当前深度为1  则处于第4行
     *
     * @param depth 深度
     * @return
     */
    private Integer getRowsIndexByDepth(int depth) {
        return titleRowIndex + MAX_TITLE_DEPTH - depth;
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
        int currentMaxDepth = StrUtil.count(title.titleName, title.PARENT_TITLE_SEPARATOR) + 1;
        if (MAX_TITLE_DEPTH != currentMaxDepth) {
            String[] split = title.titleName.split(title.PARENT_TITLE_SEPARATOR);
            String lastTitleName = split[split.length - 1];
            for (int y = 0; y < MAX_TITLE_DEPTH - currentMaxDepth; y++) {
                title.titleName = title.titleName + title.PARENT_TITLE_SEPARATOR + lastTitleName;
            }
        }
        return title;
    }

    @Override
    protected void preWrite() {
        super.preWrite();

        // fast-fail：保证sheet页初始化顺序（必须先完成上一个sheet的首次write/initTitles）
        if (sheetIndex != null && sheetIndex > 0) {
            if (!parent.metadatas.containsKey(sheetInitKey(sheetIndex - 1))) {
                throw new IllegalStateException("切换sheet页失败：sheetIndex=" + sheetIndex
                        + " 的首次写入要求 sheetIndex=" + (sheetIndex - 1)
                        + " 已经执行过一次 write() 完成初始化（允许 write(emptyList) 仅写表头）。");
            }
        }


        if (getBigExcelWriter().getSheetCount() == 1
                && getBigExcelWriter().getSheets().get(0).getPhysicalNumberOfRows() == 0
        ) {
            // 如果当前sheet页已经存在,且没有写入过, 说明是刚初始化的. 不用创建新页
            getBigExcelWriter().renameSheet(sheetName);
        } else {
            //以指定名称创建新页, 或者切换到指定页
            getBigExcelWriter().setSheet(sheetName);
        }

        // 追加写：恢复到上一次写入结束位置，避免覆盖/重复写表头
        if (isTitleInitialized()) {
            getBigExcelWriter().setCurrentRow(this.currentRow);
            return;
        }

        // 首次写：定位标题行 -> 初始化/写表头 -> 定位首行数据
        getBigExcelWriter().setCurrentRow(titleRowIndex == null ? 0 : titleRowIndex);
        this.initTitles();
        getBigExcelWriter().setCurrentRow(firstRowIndex);

        // 关键：支持首次 write(emptyList) 仅写表头后仍可追加写
        this.currentRow = getBigExcelWriter().getCurrentRow();
        parent.metadatas.put(sheetInitKey(this.sheetIndex), Boolean.TRUE);
    }

    /**
     * 以对象形式写入
     *
     * @param targetDataList 数据集
     * @return
     */
    @Override
    protected void doWrite(List<R> targetDataList) {
        /* 计算可写入的最大行数 */
        int currentRowIndex = getBigExcelWriter().getCurrentRow();
        int maxRows = latestRowIndex - currentRowIndex + 1;

        /* 如果超过限制，则截断数据 */
        List<R> dataToWrite = targetDataList;
        if (maxRows < targetDataList.size() && latestRowIndex != Integer.MAX_VALUE) {
            dataToWrite = targetDataList.subList(0, Math.max(0, maxRows));
            if (parent.debugger) {
                log.info("[Excel构建] 写入范围受限，数据被截断。原始数据行数: {}, 实际写入行数: {}", targetDataList.size(), dataToWrite.size());
            }
        }

        /* 分批写,数量上限等于滑动窗口值*/
        List<List<R>> splitDataList = ListUtil.partition(dataToWrite, BigExcelWriter.DEFAULT_WINDOW_SIZE);
        for (int i = 0; i < splitDataList.size(); i++) {

            /*写数据*/
            int startRowIndex = getBigExcelWriter().getCurrentRow();
            getBigExcelWriter().write(splitDataList.get(i).stream().map(this::changeToCellValues).toList());
            int endRowIndex = getBigExcelWriter().getCurrentRow();

            /*记录一下当当前写入的行数*/
            currentRow = getBigExcelWriter().getCurrentRow();

            /*写单元格样式*/
            for (int k = 0; k < endRowIndex - startRowIndex; k++) {
                for (int colIndex = 0; colIndex < titles.size(); colIndex++) {
                    /*写回样式*/
                    int finalI = i;
                    int finalK = k;
                    CellStyle runningTimeCellStyle = titles.get(colIndex)
                            .conditionCellStyleList
                            .stream()
                            .filter(Objects::nonNull)
                            .filter(x -> x.getCondition().test(splitDataList.get(finalI).get(finalK)))
                            .map(ConditionCellStyle::getDefaultCellStyle)
                            .map(this::getRunningTimeCellStyle)
                            .findAny()
                            .orElse(null);
                    if (runningTimeCellStyle == null) {
                        continue;
                    }
                    getBigExcelWriter().setStyle(runningTimeCellStyle, colIndex, k + startRowIndex);
                }
            }


        }

        /*将传入过来的rows按每批100次进行写入到硬盘 此时可以设置style, 底层不会报"已经写入磁盘无法编辑"的异常*/
//        List<List<R>> splitObjects = ListUtil.partition(targetDataList, BigExcelWriter.DEFAULT_WINDOW_SIZE);
//        List<List<List<Object>>> splitRows = ListUtil.partition(rows, BigExcelWriter.DEFAULT_WINDOW_SIZE);
//        for (int i = 0; i < splitRows.size(); i++) {
//            int startRowIndex = getBigExcelWriter().getCurrentRow();
//            getBigExcelWriter().write(splitRows.get(i));
//            int endRowIndex = getBigExcelWriter().getCurrentRow();
//            applyConditionStyle(splitObjects.get(i), startRowIndex, endRowIndex);
//        }


    }

    /**
     * 通过内置的样式换取为当前工作簿里面的样式
     *
     * @param style
     * @return
     */
    @Override
    protected CellStyle getRunningTimeCellStyle(ExcelCellStyle style) {
        Object cached = parent.metadatas.get(style);
        if (cached != null) {
            return (CellStyle) cached;
        }
        Workbook workbook = getWorkbook();
        CellStyle newCellStyle = StyleUtil.createDefaultCellStyle(workbook);
        style.getStyleInitializer().accept(workbook, newCellStyle);
        parent.metadatas.put(style, newCellStyle);
        return newCellStyle;
    }


    @Override
    protected void onSwitchSheet() {
        getBigExcelWriter().setSheet(sheetIndex);
    }

    @Override
    protected void stopWrite() {
        //对于hutool的实现类, stopWrite方法只需要执行一次就够了
        if (!parent.metadatas.containsKey("HUTOOL_TITLE_WRITER_CLOSED")) {
            getBigExcelWriter().close();
            parent.metadatas.put("HUTOOL_TITLE_WRITER_CLOSED", null);
        }

    }
}
