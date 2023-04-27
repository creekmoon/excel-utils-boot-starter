package cn.creekmoon.excelUtils.core;


import cn.creekmoon.excelUtils.hutool589.core.collection.ListUtil;
import cn.creekmoon.excelUtils.hutool589.core.io.FileUtil;
import cn.creekmoon.excelUtils.hutool589.core.io.IoUtil;
import cn.creekmoon.excelUtils.hutool589.core.lang.UUID;
import cn.creekmoon.excelUtils.hutool589.core.util.ArrayUtil;
import cn.creekmoon.excelUtils.hutool589.core.util.StrUtil;
import cn.creekmoon.excelUtils.hutool589.poi.excel.BigExcelWriter;
import cn.creekmoon.excelUtils.hutool589.poi.excel.ExcelUtil;
import cn.creekmoon.excelUtils.hutool589.poi.excel.style.StyleUtil;
import cn.creekmoon.excelUtils.threadPool.CleanTempFilesExecutor;
import lombok.AllArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.Collectors;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelExport<R> {
    /**
     * 表头集合
     */
    private List<Title<R>> titles = new ArrayList<>();
    /**
     * 表头集合
     */
    private HashMap<Integer, List<ConditionStyle>> colIndex2Styles = new HashMap<>();


    /* 多级表头时会用到 全局标题深度  initTitle方法会给其赋值
     *
     *   |       title               |     深度=3    rowIndex=0
     *   |   titleA    |    titleB   |     深度=2    rowIndex=1
     *   |title1|title2|title3|title4|     深度=1    rowIndex=2
     * */
    public Integer MAX_TITLE_DEPTH = null;
    /* 多级表头时会用到  深度和标题的映射关系*/
    HashMap<Integer, List<Title>> depth2Titles = new HashMap<>();

    /**
     * 当前写入行,切换sheet页时需要还原这个上下文数据
     */

    private String currentSheetName;

    /**
     * 当前写入行,切换sheet页时需要还原这个上下文数据
     */
    private int currentRow;
    /**
     * 打印调试内容
     */
    private boolean debugger = false;
    /*唯一识别名称*/
    public String taskId;
    /*自定义的名称*/
    private String excelName;
    /*写入器*/
    private BigExcelWriter bigExcelWriter;
    /*
     * sheet页和导出对象的映射关系
     * */
    private Map<String, ExcelExport> sheetName2ExcelExport = new HashMap<>();

    private ExcelExport() {

    }

    /**
     * 不指定导入类型
     *
     * @param excelName
     * @return
     */
    public static ExcelExport<Object> create(String excelName) {
        return ExcelExport.create(excelName, Object.class);
    }

    public static ExcelExport<Object> create() {
        return create("ExcelExport", Object.class);
    }

    public static <T> ExcelExport<T> create(String excelName, Class<T> c) {
        ExcelExport<T> excelExport = new ExcelExport();
        excelExport.excelName = excelName;
        excelExport.taskId = UUID.fastUUID().toString();
        return excelExport;
    }

    public static <T> ExcelExport<T> create(Class<T> c) {
        return create("ExcelExport", c);
    }


    /**
     * 添加标题
     */
    public ExcelExport<R> addTitle(String titleName, Function<R, Object> valueFunction) {
        titles.add(Title.of(titleName, valueFunction));
        return this;
    }


    /**
     * 为当前列设置一个样式
     *
     * @param condition        样式触发的条件
     * @param styleInitializer 样式初始化器
     * @return
     */
    public ExcelExport<R> setDataStyle(Predicate<R> condition, Consumer<XSSFCellStyle> styleInitializer) {
        return setDataStyle(titles.size() - 1, condition, styleInitializer);
    }

    /**
     * 添加excel样式映射,在写入的时候会读取这个映射
     *
     * @param colIndex         对应的列号
     * @param condition        样式的触发条件
     * @param styleInitializer 样式的初始化内容
     * @return
     */
    public ExcelExport<R> setDataStyle(int colIndex, Predicate<R> condition, Consumer<XSSFCellStyle> styleInitializer) {
        /*初始化样式*/
//        CellStyle newCellStyle = getBigExcelWriter().createCellStyle();
        XSSFCellStyle newCellStyle = (XSSFCellStyle) StyleUtil.createDefaultCellStyle(getBigExcelWriter().getWorkbook());
        styleInitializer.accept(newCellStyle);
        ConditionStyle conditionStyle = new ConditionStyle(condition, newCellStyle);

        /*保存映射结果*/
        if (!colIndex2Styles.containsKey(colIndex)) {
            colIndex2Styles.put(colIndex, new ArrayList<>());
        }
        colIndex2Styles.get(colIndex).add(conditionStyle);
        return this;
    }


    /**
     * 为当前列的数据设置一个样式
     *
     * @param styleInitializer 样式初始化器
     * @return
     */
    public ExcelExport<R> setDataStyle(Consumer<XSSFCellStyle> styleInitializer) {
        return this.setDataStyle(x -> true, styleInitializer);
    }


    public ExcelExport<R> debug() {
        this.debugger = true;
        return this;
    }


    /**
     * 写入对象
     *
     * @param data
     * @return
     */
    public ExcelExport<R> write(List<R> data) {
        return write(data, WriteStrategy.CONTINUE_ON_ERROR);
    }


    /**
     * 以map形式写入
     *
     * @param data key=标题 value=行内容
     * @return
     */
    protected BigExcelWriter writeByMap(List<Map> data) {
        getBigExcelWriter().write(data);
        return getBigExcelWriter();
    }


    /**
     * 以对象形式写入
     *
     * @param vos           数据集
     * @param writeStrategy 写入策略
     * @return
     */
    private ExcelExport<R> write(List<R> vos, WriteStrategy writeStrategy) {
        this.initTitles();
        List<List<Object>> rows =
                vos.stream()
                        .map(
                                vo -> {
                                    List<Object> row = new LinkedList<>();
                                    for (Title<R> title : titles) {
                                        //当前行中的某个属性值
                                        Object data = null;
                                        try {
                                            data = title.valueFunction.apply(vo);
                                        } catch (Exception exception) {
                                            if (writeStrategy == WriteStrategy.CONTINUE_ON_ERROR) {
                                                // nothing to do
                                                if (debugger) {
                                                    log.info("[Excel构建]生成Excel获取数据值时发生错误!已经忽略错误并设置为NULL值!", exception);
                                                }
                                            }
                                            if (writeStrategy == WriteStrategy.STOP_ON_ERROR) {
                                                String taskId = stopWrite();
                                                ExcelExport.cleanTempFile(taskId);
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
        //getBigExcelWriter().write(rows);

        /*将传入过来的rows按每批100次进行写入到硬盘 此时可以设置style, 底层不会报"已经写入磁盘无法编辑"的异常*/
        List<List<R>> subVos = ListUtil.partition(vos, BigExcelWriter.DEFAULT_WINDOW_SIZE);
        List<List<List<Object>>> subRows = ListUtil.partition(rows, BigExcelWriter.DEFAULT_WINDOW_SIZE);
        for (int i = 0; i < subRows.size(); i++) {
            int startRowIndex = getBigExcelWriter().getCurrentRow();
            getBigExcelWriter().write(subRows.get(i));
            int endRowIndex = getBigExcelWriter().getCurrentRow();
            setDataStyle(subVos.get(i), startRowIndex, endRowIndex);
        }


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
            for (int colIndex = 0; colIndex < titles.size(); colIndex++) {
                List<ConditionStyle> styleList = colIndex2Styles.getOrDefault(colIndex, Collections.emptyList());
                for (ConditionStyle conditionStyle : styleList) {
                    if (conditionStyle.condition.test(vo)) {
                        /*写回样式*/
                        getBigExcelWriter().setStyle(conditionStyle.style, colIndex, startRowIndex + i);
                    }
                }
            }
        }
    }

    public static void cleanTempFile(String taskId) {
        CleanTempFilesExecutor.cleanTempFileDelay(taskId);
    }

    /**
     * 初始化标题
     */
    private void initTitles() {

        /*如果已经初始化完毕 则不进行初始化*/
        if (isTitleInitialized()) {
            return;
        }

        this.MAX_TITLE_DEPTH = titles.stream()
                .map(x -> StrUtil.count(x.titleName, Title.PARENT_TITLE_SEPARATOR) + 1)
                .max(Comparator.naturalOrder())
                .orElse(1);
        if (debugger) {
            System.out.println("[Excel构建] 表头深度获取成功! 表头最大深度为" + this.MAX_TITLE_DEPTH);
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
                            if (debugger) {
                                System.out.println("[Excel构建] 插入表头[" + list.get(0).titleName + "]");
                            }
                            this.getBigExcelWriter().getOrCreateCell(startColIndex, rowsIndex).setCellValue(list.get(0).titleName);
                            this.getBigExcelWriter().getOrCreateCell(startColIndex, rowsIndex).setCellStyle(this.getBigExcelWriter().getHeadCellStyle());
                            return;
                        }
                        if (debugger) {
                            System.out.println("[Excel构建] 插入表头并横向合并[" + list.get(0).titleName + "]预计合并格数[" + (endColIndex - startColIndex) + "]");
                        }
                        this.getBigExcelWriter().merge(rowsIndex, rowsIndex, startColIndex, endColIndex, list.get(0).titleName, true);
                        if (debugger) {
                            System.out.println("[Excel构建] 横向合并表头[" + list.get(0).titleName + "]完毕");
                        }
                    });
            /* hutool excel 写入下移一行*/
            this.getBigExcelWriter().setCurrentRow(this.getBigExcelWriter().getCurrentRow() + 1);
        }

        /*纵向合并title*/
        for (int colIndex = 0; colIndex < titles.size(); colIndex++) {
            Title<R> title = titles.get(colIndex);
            int sameCount = 0; //重复的数量
            for (Integer depth = 1; depth <= MAX_TITLE_DEPTH; depth++) {
                if (title.parentTitle != null && title.titleName.equals(title.parentTitle.titleName)) {
                    /*发现重复的单元格,不会马上合并 因为可能有更多重复的*/
                    sameCount++;
                } else if (sameCount != 0) {
                    if (debugger) {
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
     * 停止写入
     *
     * @return taskId
     */
    public String stopWrite() {
        getBigExcelWriter().close();
        return taskId;
    }


    /**
     * 默认设置列宽 : 简单粗暴将前500列都设置宽度20
     */
    public void setColumnWidthDefault() {
        this.setColumnWidth(500, 20);
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

    /**
     * 响应请求 返回结果
     *
     * @param taskId
     * @param responseExcelName
     * @param response
     * @throws IOException
     */
    public static void response(String taskId, String responseExcelName, HttpServletResponse response) throws IOException {
        try {
            if (response != null) {
                responseByFilePath(PathFinder.getAbsoluteFilePath(taskId), responseExcelName, response);
            }
        } finally {
            cleanTempFile(taskId);
        }
    }


    /**
     * 返回Excel
     *
     * @param filePath          本地文件路径
     * @param responseExcelName 声明的文件名称,前端能看到 可以自己乱填
     * @param response          servlet请求
     * @throws IOException
     */
    /*回应请求*/
    private static void responseByFilePath(String filePath, String responseExcelName, HttpServletResponse response) throws IOException {
        /*现代浏览器标准, RFC5987标准协议 显式指定文件名的编码格式为UTF-8 但是这样swagger-ui不支持回显 比较坑*/
//        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//        response.setHeader("Content-Disposition", "attachment;filename*=UTF-8''" + URLEncoder.encode(responseExcelName + ".xlsx", "UTF-8"));

        /*旧版协议 不能使用中文文件名 但是兼容所有浏览器*/
//        response.setContentType("application/octet-stream");
//        response.setHeader("Content-Disposition", "attachment;filename=" + responseExcelName + ".xlsx");

        /*当前的兼容做法, 使用旧版协议, 但是用UTF-8编码文件名. 这样一来支持旧版浏览器下载文件(但文件名乱码), 同时现代浏览器能下载也能会自动解析成中文文件名*/
        response.setContentType("application/octet-stream");
        responseExcelName = URLEncoder.encode(responseExcelName + ".xlsx", "UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + responseExcelName);

        /*使用流将文件传输回去*/
        ServletOutputStream out = null;
        InputStream fileInputStream = null;
        try {
            out = response.getOutputStream();
            fileInputStream = FileUtil.getInputStream(filePath);
            byte[] b = new byte[4096];  //创建数据缓冲区  通常网络2-8K 磁盘32K-64K
            int length;
            while ((length = fileInputStream.read(b)) > 0) {
                out.write(b, 0, length);
            }
            out.flush();
            out.close();
        } finally {
            IoUtil.close(fileInputStream);
            IoUtil.close(out);
        }

    }

    /**
     * 响应并清除文件
     *
     * @param response
     * @throws IOException
     */
    public void response(HttpServletResponse response) throws IOException {
        String taskId = this.stopWrite();
        ExcelExport.response(taskId, excelName, response);
    }


    /**
     * 内部操作类,但是暴露出来了,希望最好不要用这个方法
     *
     * @return
     */
    public BigExcelWriter getBigExcelWriter() {
        if (bigExcelWriter == null) {
            bigExcelWriter = ExcelUtil.getBigWriter(PathFinder.getAbsoluteFilePath(taskId), currentSheetName);
        }
        return bigExcelWriter;
    }


    /**
     * 获取行坐标   如果最大深度为4  当前深度为1  则处于第4行
     *
     * @param depth 深度
     * @return
     */
    private Integer getRowsIndexByDepth(int depth) {
        return MAX_TITLE_DEPTH - depth;
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
        if (MAX_TITLE_DEPTH != currentMaxDepth) {
            String[] split = title.titleName.split(Title.PARENT_TITLE_SEPARATOR);
            String lastTitleName = split[split.length - 1];
            for (int y = 0; y < MAX_TITLE_DEPTH - currentMaxDepth; y++) {
                title.titleName = title.titleName + Title.PARENT_TITLE_SEPARATOR + lastTitleName;
            }
        }
        return title;
    }

    /**
     * 标题类  是个单向链表 会指向自己的父表头
     *
     * @param <T>
     */
    public static class Title<T> {
        public Title parentTitle;
        public String titleName;
        public Function<T, Object> valueFunction;
        /* 父标题分隔符*/
        public static String PARENT_TITLE_SEPARATOR = "::";
        /* 列坐标 开始于第几列  */
        public int startColIndex;
        /* 列坐标 结束于第几列 */
        public int endColIndex;

        public Title(String titleName, Function<T, Object> valueFunction) {
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
        private HashMap<Integer, Title> convert2ChainTitle(int currentColIndex) {
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


    /**
     * 切换到新的标签页
     */
    public <T> ExcelExport<T> switchSheet(String sheetName, Class<T> newDataClass) {

        /*如果当前sheet是空内容, 则放弃当前的sheet*/
        if (!isTitleInitialized() && titles.isEmpty()) {
            this.getBigExcelWriter().renameSheet(sheetName);
        }

        /*保存当前的写入记录*/
        this.currentRow = Optional.ofNullable(bigExcelWriter)
                .map(BigExcelWriter::getCurrentRow)
                .orElse(0);

        /*尝试获取已存在的sheet*/
        ExcelExport cached = sheetName2ExcelExport.get(sheetName);
        if (cached != null) {
            cached.getBigExcelWriter().setSheet(cached.currentSheetName);
            cached.getBigExcelWriter().setCurrentRow(cached.currentRow);
            return cached;
        }

        /*创建一个新的导出对象, 并关联一个新的sheet*/
        ExcelExport<T> newExcelExport = ExcelExport.create(this.excelName, newDataClass);
        newExcelExport.currentSheetName = sheetName;
        newExcelExport.taskId = taskId;
        newExcelExport.debugger = debugger;
        newExcelExport.bigExcelWriter = bigExcelWriter;
        newExcelExport.sheetName2ExcelExport = sheetName2ExcelExport;
        newExcelExport.getBigExcelWriter().setSheet(sheetName);
        this.sheetName2ExcelExport.put(sheetName, newExcelExport);
        return newExcelExport;
    }


    /**
     * 写入策略
     */
    public enum WriteStrategy {
        /*忽略取值异常 通常是多级属性空指针造成的 如果取不到值直接置为NULL*/
        CONTINUE_ON_ERROR,
        /*遇到任何失败的情况则停止*/
        STOP_ON_ERROR;
    }

    /*条件样式*/
    @AllArgsConstructor
    private class ConditionStyle {
        Predicate<R> condition;
        CellStyle style;
    }
}
