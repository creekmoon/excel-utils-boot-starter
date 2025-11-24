package cn.creekmoon.excel.util;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.W.ExcelExport;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IoUtil;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.springframework.scheduling.concurrent.CustomizableThreadFactory;
import org.springframework.util.ResourceUtils;

import java.io.*;
import java.net.URISyntaxException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Paths;
import java.util.concurrent.ScheduledThreadPoolExecutor;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

/**
 * 获取路径
 */
@Slf4j
public class ExcelFileUtils {

    /*临时文件目录*/
    private static volatile String applicationParentFilePath;

    private static final ThreadFactory threadFactory = new CustomizableThreadFactory("excel-clean-thread");

    /*一个延迟任务线程池*/
    private static final ScheduledThreadPoolExecutor threadPoolExecutor = new ScheduledThreadPoolExecutor(1, threadFactory, new ThreadPoolExecutor.AbortPolicy());


    /**
     * 生成Xlsx文件绝对路径 默认路径在当前jar包的同级目录
     *
     * @param uniqueId 唯一名称,写文件完成后会获得
     * @return
     */
    public static String generateXlsxAbsoluteFilePath(String uniqueId) {
        return getCustomAbsoluteFilePath(uniqueId + ".xlsx");
    }

    /**
     * 获取获取一个文件路径 默认路径在当前jar包的同级目录
     *
     * @param fileName 文件名称 需要自己保证不重复
     * @return
     */
    public static String getCustomAbsoluteFilePath(String fileName) {
//这里凌晨跨天的话会有问题 不过我不想管了... hhh
//        return getApplicationParentFilePath() + File.separator + "temp_files" + File.separator + DateUtil.format(new Date(), "yyyy-MM-dd") + File.separator + fileName;
        return getApplicationParentFilePath() + File.separator + "temp_files" + File.separator + fileName;
    }


    protected static void response(InputStream fileInputStream, String fileName, HttpServletResponse response) throws IOException {
        /*现代浏览器标准, RFC5987标准协议 显式指定文件名的编码格式为UTF-8 但是这样swagger-ui不支持回显 比较坑*/
//        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//        response.setHeader("Content-Disposition", "attachment;filename*=UTF-8''" + URLEncoder.encode(fileName + ".xlsx", "UTF-8"));

        /*旧版协议 不能使用中文文件名 但是兼容所有浏览器*/
//        response.setContentType("application/octet-stream");
//        response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");

        /*当前的兼容做法, 使用旧版协议, 但是用UTF-8编码文件名. 这样一来支持旧版浏览器下载文件(但文件名乱码), 同时现代浏览器能下载也能会自动解析成中文文件名*/
        response.setContentType("application/octet-stream");
        fileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8);
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName);

        /*使用流将文件传输回去*/
        try (ServletOutputStream out = response.getOutputStream()) {
            byte[] b = new byte[4096];  //创建数据缓冲区  通常网络2-8K 磁盘32K-64K
            int length;
            while ((length = fileInputStream.read(b)) > 0) {
                out.write(b, 0, length);
            }
            out.flush();
        } finally {
            IoUtil.close(fileInputStream);
        }
    }

    /**
     * 返回Excel
     *
     * @param file     本地文件
     * @param fileName 声明的文件名称,前端能看到 可以自己乱填
     * @param response servlet请求
     * @throws IOException
     */
    /*回应请求*/
    public static void response(File file, String fileName, HttpServletResponse response) throws IOException {
        response(file.getAbsolutePath(), fileName, response);
    }


    /**
     * 返回Excel
     *
     * @param filePath 本地文件路径
     * @param fileName 声明的文件名称,前端能看到 可以自己乱填
     * @param response servlet请求
     * @throws IOException
     */
    /*回应请求*/
    public static void response(String filePath, String fileName, HttpServletResponse response) throws IOException {
        try (BufferedInputStream inputStream = FileUtil.getInputStream(filePath)) {
            response(inputStream, fileName, response);
        }
    }

    public static void cleanTempFileByPathDelay(File file) {
        cleanTempFileByPathDelay(file.getAbsolutePath());
    }

    public static void cleanTempFileByPathDelay(String filePath) {
        cleanTempFileByPathDelay(filePath, ExcelUtilsConfig.TEMP_FILE_LIFE_MINUTES);
    }

    public static void cleanTempFileByPathDelay(String filePath, Integer fileLifeMinutes) {
        threadPoolExecutor.schedule(() -> {
            cleanTempFileByPathNow(filePath);
        }, fileLifeMinutes != null ? fileLifeMinutes : ExcelUtilsConfig.TEMP_FILE_LIFE_MINUTES, TimeUnit.MINUTES);
    }


    public static void cleanTempFileByPathNow(String filePath) {
        try {
            if (!FileUtil.del(filePath)) {
                log.warn("清理临时文件失败! 路径:" + filePath);
            }
            log.debug("清理临时文件成功!路径:" + filePath);
        } catch (Exception e) {
            log.warn("清理临时文件失败! 路径:" + filePath);
        }
    }


    /*获取项目路径*/
    @SneakyThrows
    protected static String getApplicationParentFilePath() {
        if (applicationParentFilePath == null) {
            synchronized (ExcelExport.class) {
                if (applicationParentFilePath == null) {
                    applicationParentFilePath = System.getProperty("user.dir");
                }
            }
        }
        return applicationParentFilePath;
    }
}
