package cn.jy.excelUtils.threadPool;

import cn.jy.excelUtils.core.AsyncTaskState;
import lombok.extern.slf4j.Slf4j;
import org.springframework.scheduling.concurrent.CustomizableThreadFactory;

import java.util.concurrent.*;
import java.util.function.Consumer;

/**
 * 专门用于执行回调任务的线程池(监控SAXReader的执行情况)
 *
 * @author jy
 */
@Slf4j
public class AsyncStateCallbackExecutor {
    private static ThreadFactory namedThreadFactory = new CustomizableThreadFactory("excel-callback-thread");
    /**
     * 刷新间隔时间 单位秒
     */
    public static int REFRESH_MILLISECONDS;
    //两个线程的固定线程池
    private static ScheduledThreadPoolExecutor threadPoolExecutor;
    /**
     * 记录线程池的执行情况
     */
    private static ConcurrentHashMap<String, ScheduledFuture> taskId2TaskFuture = new ConcurrentHashMap<>();
    /**
     * 记录执行状态对象
     */
    private static ConcurrentHashMap<String, AsyncTaskState> taskId2TaskState = new ConcurrentHashMap<>();
    /**
     * 记录用户定义的回调函数
     */
    private static ConcurrentHashMap<String, Consumer<AsyncTaskState>> taskId2CallbackFunc = new ConcurrentHashMap<>();

    public static void init(int corePoolSize) {
        if (threadPoolExecutor != null) {
            threadPoolExecutor.shutdown();
        }
        threadPoolExecutor = new ScheduledThreadPoolExecutor(corePoolSize, namedThreadFactory);

    }

    /**
     * 创建一个回调任务
     *
     * @param taskId
     * @param asyncCallback
     * @return
     */
    public static AsyncTaskState createAsyncTaskState(String taskId, Consumer<AsyncTaskState> asyncCallback) {
        AsyncTaskState asyncTaskState = new AsyncTaskState();
        asyncTaskState.setTaskId(taskId);
        asyncTaskState.setWaiting(true);
        Runnable runnable = () -> {
            if (asyncTaskState.isCompleted()) {
                //通常不会走到这一步 因为任务标记完成之后,不会在线程池内存活了
                log.error("任务已经完成,不应该再次执行! 如果此错误连续发生,请检查代码逻辑!");
                return;
            }
            if (asyncTaskState.getSuccessRowIndex() > 0) {
                asyncTaskState.setWaiting(false);
            }
            asyncCallback.accept(asyncTaskState);
        };

        /*执行定时回调*/
        ScheduledFuture taskFuture = threadPoolExecutor.scheduleAtFixedRate(runnable, 0, REFRESH_MILLISECONDS, TimeUnit.MILLISECONDS);

        /*记录状态*/
        taskId2TaskState.put(taskId, asyncTaskState);
        taskId2TaskFuture.put(taskId, taskFuture);
        taskId2CallbackFunc.put(taskId, asyncCallback);
        return asyncTaskState;
    }

    /**
     * 标记回调任务完成
     *
     * @param taskId
     */
    public static void completedAsyncTaskState(String taskId) {
        //停止在线程池中回调
        ScheduledFuture scheduledFuture = taskId2TaskFuture.get(taskId);
        if (scheduledFuture != null) {
            scheduledFuture.cancel(false);
            taskId2TaskFuture.remove(taskId);
        }
        //设置任务状态为已完成
        AsyncTaskState asyncTaskState = taskId2TaskState.get(taskId);
        if (asyncTaskState != null) {
            asyncTaskState.setCompleted(true);
            taskId2TaskState.remove(taskId);
        }
        //执行最后一次回调 这样能保证完成任务时,至少能够执行一次回调函数
        Consumer<AsyncTaskState> asyncTaskStateConsumer = taskId2CallbackFunc.get(taskId);
        if (asyncTaskStateConsumer != null) {
            asyncTaskStateConsumer.accept(asyncTaskState);
            taskId2CallbackFunc.remove(taskId);
        }
    }
}
