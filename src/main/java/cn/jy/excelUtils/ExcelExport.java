package cn.jy.excelUtils;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.lang.UUID;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;
import org.springframework.util.ResourceUtils;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
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

    /*唯一识别名称*/
    private String uniqueName;
    /*自定义的名称*/
    private String excelName;
    /*写入器*/
    private BigExcelWriter bigExcelWriter;

    private ExcelExport() {

    }

    /**
     *  不指定导入类型
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
//        List<R> apply = dataSource.apply(pageQuery);
//        for (int i = 1; i < 1001; i++) {
//            pageable.setPageNo(i);
//            this.write(apply);
//            if (pageable.getTotalPage() == null || pageable.getTotalPage() <= i) {
//                break;
//            }
//        }
//        this.send(response);
//    }

    /**
     * 写入对象
     * @param data
     * @return
     */
    public BigExcelWriter write(List<R> data) {
        return write(data, titles);
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
     * @param vos    数据集
     * @param titles 标题映射关系
     * @param <T>
     * @return
     */
    public <T> BigExcelWriter write(List<T> vos, List<Title<T>> titles) {
        List<HashMap<String, Object>> rows =
                vos.stream().map(
                                vo -> {
                                    HashMap<String, Object> row = new HashMap<>();
                                    for (Title<T> title : titles) {
                                        row.put(title.titleName, title.valueFunction.apply(vo));
                                    }
                                    return row;
                                })
                        .collect(Collectors.toList());
        getBigExcelWriter().write(rows);
        return getBigExcelWriter();

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
    private void autoSetColumnWidth(){
        for (int i = 0; i <100 ; i++) {
            try {
                getBigExcelWriter().setColumnWidth(i, 20);
            }catch (Exception ignored){ }
        }
    }

    /*回应请求 这个步骤会将磁盘文件清空*/
    public static void send(HttpServletResponse response, String uniqueName, String excelName) throws IOException {
        /*实际上是xlsx格式的文件  但以xls格式进行发送,好像也没什么问题*/
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + excelName + ".xls");

        /*使用流将文件传输回去v*/
        ServletOutputStream out = response.getOutputStream();
        FileInputStream fileInputStream = new FileInputStream(getAbsoluteFilePath(uniqueName));
        byte[] b = new byte[1024];  //创建数据缓冲区
        int length;
        while ((length = fileInputStream.read(b)) > 0) {
            out.write(b, 0, length);
        }
        out.flush();
        out.close();
        IoUtil.close(fileInputStream);
        IoUtil.close(out);
        /*清除临时文件*/
        cleanTempFile(uniqueName);
    }

    /*将文件响应给请求*/
    public void send(HttpServletResponse response) throws IOException {
        String uniqueName = this.stopWrite();
        ExcelExport.send(response, uniqueName, excelName);
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
        File file = new File(getAbsoluteFilePath(uniqueName));
        if (file.exists() && !file.delete()) {
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
     * 标题类
     *
     * @param <T>
     */
    public static class Title<T> {
        public String titleName;
        public Function<T, Object> valueFunction;

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
    }
}
