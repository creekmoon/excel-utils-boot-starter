package cn.jy.excelUtils;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.lang.UUID;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;
import org.springframework.util.ResourceUtils;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;
import java.util.concurrent.locks.ReentrantLock;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * @author JY
 * @date 2022-01-05
 */
public class ExcelExport<R> {
    /*临时文件目录*/
    private static volatile String applicationParentFilePath;
    private static final ReentrantLock reentrantLock = new ReentrantLock();
    private List<Title<R>> titles = new ArrayList<>();
    /**
     * 打印调试内容
     */
    private boolean showDebuggerDetail = false;
    /* 多级表头时会用到 全局标题深度  initTitle方法会给其赋值
     *
     *   |       title               |     深度=3    rowIndex=0
     *   |   titleA    |    titleB   |     深度=2    rowIndex=1
     *   |title1|title2|title3|title4|     深度=1    rowIndex=2
     * */
    public Integer MAX_TITLE_DEPTH = null;
    /* 多级表头时会用到    深度 和 标题集合的映射关系*/
    HashMap<Integer, List<Title>> depth2Titles = new HashMap<>();

    /*唯一识别名称*/
    private String uniqueName;
    /*自定义的名称*/
    private String excelName;
    /*写入器*/
    private BigExcelWriter bigExcelWriter;

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

    public static <T> ExcelExport<T> create(String excelName, Class<T> c) {
        ExcelExport<T> excelExport = new ExcelExport();
        excelExport.excelName = excelName;
        excelExport.uniqueName = UUID.fastUUID().toString();
        return excelExport;
    }

    /**
     * 添加标题
     */
    public ExcelExport<R> addTitle(String titleName, Function<R, Object> valueFunction) {
        titles.add(Title.of(titleName, valueFunction));
        return this;
    }

    public void debug() {
        this.showDebuggerDetail = true;
    }
    /**
     * 语法糖 自动进行分页并导出 最高支持100W行
     *
     * @param response   spring-HttpServletResponse
     * @param pageQuery  继承了PageQuery的类 通常是PO
     * @param dataSource 数据源 通常是queryList
     * @param <T>
     * @throws IOException
     */
//    public <T> void wtxSimpleExport(HttpServletResponse response, T pageQuery, Function<T, List<R>> dataSource) throws IOException {
//        PageQuery pageable = (PageQuery) pageQuery;
//        pageable.setPageSize(1000);
//        for (int i = 1; i < 1001; i++) {
//            pageable.setPageNo(i);
//            List<R> apply = dataSource.apply(pageQuery);
//            this.writeAndIgnoreValueGetterException(apply);
//            if (pageable.getTotalPage() == null || pageable.getTotalPage() <= i) {
//                break;
//            }
//        }
//        this.response(response);
//    }

    /**
     * 写入对象
     *
     * @param data
     * @return
     */
    public BigExcelWriter write(List<R> data) {
        return write(data, titles, false);
    }

    /**
     * 写入对象 并忽略未捕获的getter异常
     *
     * @param data
     * @return
     */
    public BigExcelWriter writeAndIgnoreValueGetterException(List<R> data) {
        return write(data, titles, true);
    }

    /**
     * 以map形式写入
     *
     * @param data key=标题 value=行内容
     * @return
     */
    public BigExcelWriter writeByMap(List<Map> data) {
        getBigExcelWriter().write(data);
        return getBigExcelWriter();
    }


    /**
     * 以对象形式写入
     *
     * @param vos                               数据集
     * @param titles                            标题映射关系
     * @param ignoreValueGetterUnCatchException 忽略值的getter方法异常 通常是多层get导致空指针
     * @param <T>
     * @return
     */
    private <T> BigExcelWriter write(List<T> vos, List<Title<T>> titles, boolean ignoreValueGetterUnCatchException) {
        this.initTitles();
        List<List<Object>> rows =
                vos.stream()
                        .map(
                                vo -> {
                                    List<Object> row = new LinkedList<>();
                                    for (Title<T> title : titles) {
                                        Object apply = null;
                                        if (ignoreValueGetterUnCatchException) {
                                            try {
                                                apply = title.valueFunction.apply(vo);
                                            } catch (Exception ignored) {
                                            }
                                        } else {
                                            apply = title.valueFunction.apply(vo);
                                        }
                                        row.add(apply);
                                    }
                                    return row;
                                })
                        .collect(Collectors.toList());
        getBigExcelWriter().write(rows);
        return getBigExcelWriter();
    }

    /**
     * 初始化标题
     */
    public void initTitles() {

        /*如果已经初始化完毕 则不进行初始化*/
        if (this.MAX_TITLE_DEPTH != null) {
            return;
        }

        this.MAX_TITLE_DEPTH = titles.stream()
                .map(x -> StrUtil.count(x.titleName, Title.PARENT_TITLE_SEPARATOR) + 1)
                .max(Comparator.naturalOrder())
                .orElse(1);
        if (showDebuggerDetail) {
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
                            if (showDebuggerDetail) {
                                System.out.println("[Excel构建] 插入表头" + list.get(0).titleName);
                            }
                            this.getBigExcelWriter().getOrCreateCell(startColIndex, rowsIndex).setCellValue(list.get(0).titleName);
                            this.getBigExcelWriter().getOrCreateCell(startColIndex, rowsIndex).setCellStyle(this.getBigExcelWriter().getHeadCellStyle());
                            return;
                        }
                        if (showDebuggerDetail) {
                            System.out.println("[Excel构建] 插入表头并横向合并" + list.get(0).titleName + "  预计合并格数" + (endColIndex - startColIndex));
                        }
                        this.getBigExcelWriter().merge(rowsIndex, rowsIndex, startColIndex, endColIndex, list.get(0).titleName, true);
                        if (showDebuggerDetail) {
                            System.out.println("[Excel构建] 横向合并表头" + list.get(0).titleName + "完毕");
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
                    if (showDebuggerDetail) {
                        System.out.println("[Excel构建] 表头纵向合并" + title.titleName + "  预计合并格数" + sameCount);
                    }
                    /*合并单元格*/
                    this.getBigExcelWriter().merge(getRowsIndexByDepth(depth), getRowsIndexByDepth(depth) + sameCount, colIndex, colIndex, title.titleName, true);
                    sameCount = 0;
                }
                title = title.parentTitle;
            }
        }

    }

    /*保存文件*/
    private String stopWrite() {
        this.autoSetColumnWidth();
        getBigExcelWriter().close();
        return uniqueName;
    }


    /**
     * 自动设置列宽 : 简单粗暴将前100列都设置宽度
     */
    private void autoSetColumnWidth() {
        for (int i = 0; i < 100; i++) {
            try {
                getBigExcelWriter().setColumnWidth(i, 20);
            } catch (Exception ignored) {
            }
        }
    }

    /*回应请求 这个步骤会将磁盘文件清空*/
    public static void response(HttpServletResponse response, String uniqueName, String excelName) throws IOException {
        /*实际上是xlsx格式的文件  但以xls格式进行发送,好像也没什么问题*/
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + excelName + ".xls");

        /*使用流将文件传输回去*/
        ServletOutputStream out = null;
        InputStream fileInputStream = null;
        try {
            out = response.getOutputStream();
            fileInputStream = FileUtil.getInputStream(getAbsoluteFilePath(uniqueName));
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
            /*清除临时文件*/
            cleanTempFile(uniqueName);
        }

    }

    /*将文件响应给请求*/
    public void response(HttpServletResponse response) throws IOException {
        String uniqueName = this.stopWrite();
        ExcelExport.response(response, uniqueName, excelName);
    }


    private BigExcelWriter getBigExcelWriter() {
        if (bigExcelWriter == null) {
            bigExcelWriter = ExcelUtil.getBigWriter(getAbsoluteFilePath(uniqueName));
        }
        return bigExcelWriter;
    }

    /*获取项目路径*/
    private static String getApplicationParentFilePath() {
        if (applicationParentFilePath == null) {
            reentrantLock.lock();
            try {
                if (applicationParentFilePath == null) {
                    applicationParentFilePath = new File(ResourceUtils.getURL("classpath:").getPath()).getParentFile().getParentFile().getParent();
                }
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } finally {
                reentrantLock.unlock();
            }
        }
        return applicationParentFilePath;
    }

    /**
     * 清理临时文件
     *
     * @param uniqueName
     * @throws IOException
     */
    private static void cleanTempFile(String uniqueName) throws IOException {
        if (FileUtil.del(getAbsoluteFilePath(uniqueName))) {
            throw new IOException("清理临时文件失败! 路径:" + getAbsoluteFilePath(uniqueName));
        }
    }

    /**
     * 获取文件绝对路径
     *
     * @param uniqueName
     * @return
     */
    private static String getAbsoluteFilePath(String uniqueName) {
        return getApplicationParentFilePath() + File.separator + "temp_files" + File.separator + DateUtil.format(new Date(), "yyyy-MM-dd") + File.separator + uniqueName + ".xlsx";
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
}
