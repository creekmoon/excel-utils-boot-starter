package cn.creekmoon.excelUtils.core;


import cn.creekmoon.excelUtils.threadPool.CleanTempFilesExecutor;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.lang.UUID;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.jetbrains.annotations.NotNull;

import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Map;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelExport {
    /**
     * 打印调试内容
     */
    protected boolean debugger = false;
    /*唯一识别名称*/
    public String taskId = UUID.fastUUID().toString();
    /*自定义的名称*/
    public String excelName;
    /*写入器*/
    private BigExcelWriter bigExcelWriter;
    /*
     * sheet页和导出对象的映射关系
     * */
    private Map<String, SheetWriter> sheetName2SheetWriter = new HashMap<>();

    private ExcelExport() {

    }

    public static <T> SheetWriter<T> create(Class<T> c) {
        ExcelExport excelExport = create();
        return excelExport.switchSheet(c);
    }

    public static ExcelExport create() {
        ExcelExport excelExport = new ExcelExport();
        excelExport.excelName = ExcelConstants.excelNameGenerator.get();
        return excelExport;
    }

    /**
     * 生成sheetName
     *
     * @param sheetIndex sheet页下标
     * @return
     */
    @NotNull
    protected static String generateSheetNameByIndex(Integer sheetIndex) {
        return "Sheet" + (sheetIndex + 1);
    }

    /**
     * 切换到新的标签页
     */
    public <T> SheetWriter<T> switchSheet(String sheetName, Class<T> newDataClass) {
        /*第一次切换sheet页是重命名当前sheet*/
        if (sheetName2SheetWriter.isEmpty()) {
            getBigExcelWriter().renameSheet(sheetName);
        }
        /*后续切换sheet页是新增sheet*/
        if (!sheetName2SheetWriter.isEmpty()) {
            getBigExcelWriter().setSheet(sheetName);
        }
        return sheetName2SheetWriter.computeIfAbsent(sheetName, s -> {
            return new SheetWriter<>(this, new SheetWriterContext(sheetName));
        });

    }

    /**
     * 切换到新的标签页
     */
    public <T> SheetWriter<T> switchSheet(Class<T> newDataClass) {
        int indexSeq = 0;
        while (sheetName2SheetWriter.containsKey(generateSheetNameByIndex(indexSeq))) {
            indexSeq++;
        }
        return switchSheet(generateSheetNameByIndex(indexSeq), newDataClass);
    }

    public ExcelExport debug() {
        this.debugger = true;
        return this;
    }


    public static void cleanTempFileDelay(String taskId) {
        CleanTempFilesExecutor.cleanTempFileDelay(taskId);
    }

    public void cleanTempFileDelay() {
        CleanTempFilesExecutor.cleanTempFileDelay(taskId);
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
            cleanTempFileDelay(taskId);
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
            bigExcelWriter = ExcelUtil.getBigWriter(PathFinder.getAbsoluteFilePath(taskId));
        }
        return bigExcelWriter;
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

}
