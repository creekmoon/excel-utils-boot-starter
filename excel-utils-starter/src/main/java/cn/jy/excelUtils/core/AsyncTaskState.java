package cn.jy.excelUtils.core;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;

/**
 * 异步状态
 */
@Data
public class AsyncTaskState {
    /*本次任务id*/
    String taskId;
    /*成功的行号*/
    int successRowIndex = 0;
    /*是否结束 不能用这个状态区分成功还是失败,只代表任务在线程池中已经执行完了*/
    boolean completed = false;
    /*是否处于等待状态(如果超过了并发导入数量 则会等待其他任务先完成)*/
    boolean waiting = true;
    /*key=行号 value=错误信息*/
    Map<Long, String> errorReport = new HashMap<>();


}
