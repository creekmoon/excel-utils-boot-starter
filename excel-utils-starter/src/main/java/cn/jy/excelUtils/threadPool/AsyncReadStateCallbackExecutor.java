package cn.jy.excelUtils.threadPool;

import org.springframework.scheduling.concurrent.CustomizableThreadFactory;

import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.ScheduledThreadPoolExecutor;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.TimeUnit;

/**
 * 专门用于执行回调任务的线程池(监控SAXReader的执行情况)
 *
 * @author jy
 */
public class AsyncReadStateCallbackExecutor {
    private static ThreadFactory namedThreadFactory = new CustomizableThreadFactory("excel-callback-thread");
    public static int taskDelay;
    public static boolean initialized = false;
    //两个线程的固定线程池
    private static ScheduledThreadPoolExecutor threadPoolExecutor;


    public static void init(int corePoolSize, int taskDelay) {
        if (initialized) {
            return;
        }
        threadPoolExecutor = new ScheduledThreadPoolExecutor(corePoolSize, namedThreadFactory);
        AsyncReadStateCallbackExecutor.taskDelay = taskDelay;
        initialized = true;
    }

    public static ScheduledFuture run(Runnable runnable) {
        return threadPoolExecutor.scheduleAtFixedRate(runnable, 0, taskDelay, TimeUnit.MILLISECONDS);
    }

}
