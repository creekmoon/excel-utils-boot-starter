package cn.creekmoon.excelUtils.threadPool;

import cn.creekmoon.excelUtils.config.ExcelUtilsConfig;
import cn.creekmoon.excelUtils.core.PathFinder;
import cn.creekmoon.excelUtils.hutool589.core.io.FileUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.scheduling.concurrent.CustomizableThreadFactory;

import java.io.IOException;
import java.util.concurrent.ScheduledThreadPoolExecutor;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;


/**
 * 线程池 专门用于清理临时文件
 */
@Slf4j
public class CleanTempFilesExecutor {

    private static final ThreadFactory threadFactory = new CustomizableThreadFactory("excel-clean-thread");
    ;
    /*一个延迟任务线程池*/
    private static ScheduledThreadPoolExecutor threadPoolExecutor = new ScheduledThreadPoolExecutor(1, threadFactory, new ThreadPoolExecutor.AbortPolicy());


    public static void init() {
        int corePoolSize = 1;
        if (threadPoolExecutor != null) {
            threadPoolExecutor.shutdown();
        }
        threadPoolExecutor = new ScheduledThreadPoolExecutor(corePoolSize, threadFactory);
    }

    /**
     * 清理临时文件
     *
     * @param taskId
     * @throws IOException
     */
    public static void cleanTempFileDelay(String taskId) {
        threadPoolExecutor.schedule(() -> {
            cleanTeamFileNow(taskId);
        }, ExcelUtilsConfig.TEMP_FILE_LIFE_MINUTES, TimeUnit.MINUTES);
    }

    public static void cleanTeamFileNow(String taskId) {
        try {
            if (!FileUtil.del(PathFinder.getAbsoluteFilePath(taskId))) {
                log.warn("[Excel导出]清理临时文件失败! 路径:" + PathFinder.getAbsoluteFilePath(taskId));
            }
            log.debug("[Excel导出]清理临时文件成功!路径:" + PathFinder.getAbsoluteFilePath(taskId));
        } catch (Exception e) {
            log.warn("[Excel导出]清理临时文件失败! 路径:" + PathFinder.getAbsoluteFilePath(taskId));
        }
    }
}
