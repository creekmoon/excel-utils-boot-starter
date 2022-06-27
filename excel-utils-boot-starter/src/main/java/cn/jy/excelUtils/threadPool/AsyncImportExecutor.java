package cn.jy.excelUtils.threadPool;

import org.springframework.scheduling.concurrent.CustomizableThreadFactory;

import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadFactory;

/**
 * 专门用于执行异步读Excel任务的线程池
 *
 * @author jy
 */
public class AsyncImportExecutor {
    private static ThreadFactory namedThreadFactory = new CustomizableThreadFactory("excel-read-thread");
    //两个线程的固定线程池
    private static ExecutorService threadPoolExecutor = Executors.newCachedThreadPool(namedThreadFactory);

    public static void run(Runnable runnable) {
        threadPoolExecutor.submit(runnable);
    }

}
