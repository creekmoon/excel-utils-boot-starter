package cn.jy.excelUtils.core;

import lombok.Data;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 异步状态
 */
@Data
public class AsyncTaskState {
    /**
     * 保留最近的15条异常信息
     */
    private static final int LAST_ERROR_COUNT = 15;
    /*本次任务id*/
    String taskId;
    /*尝试的次数*/
    int tryRowCount = 0;
    /*成功的次数*/
    int successRowCount = 0;
    /*是否结束 不能用这个状态区分成功还是失败,只代表任务在线程池中已经执行完了*/
    boolean completed = false;
    /*是否处于等待状态(如果超过了并发导入数量 则会等待其他任务先完成)*/
    boolean waiting = true;
    /*key=行号 value=错误信息*/
    Map<Long, String> errorReport = new LinkedHashMap<>(32);


    public void addErrorMsg(Long index, String errorMsg) {
        errorReport.put(index, errorMsg);
        //删除最旧的一条错误信息
        if (errorReport.size() > LAST_ERROR_COUNT) {
            errorReport.remove(errorReport.keySet().iterator().next());
        }

    }

}
