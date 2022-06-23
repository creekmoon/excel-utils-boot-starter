package cn.jy.excelUtils.core;

import cn.hutool.core.lang.UUID;
import lombok.Data;

import java.util.HashMap;
import java.util.Map;

/**
 * 异步读取状态
 */
@Data
public class AsyncReadState {
    /*本次导入id*/
    String taskId = UUID.fastUUID().toString();
    /*成功读取并消费的行号*/
    int successRowIndex = 0;
    /*是否读取完成*/
    boolean completed = false;
    /*是否处于等待状态(如果超过了并发导入数量 则会等待其他任务先完成)*/
    boolean waiting = true;
    /*key=行号 value=错误信息*/
    Map<Long, String> errorReport = new HashMap<>();
}
