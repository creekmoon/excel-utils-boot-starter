package cn.creekmoon.excel.core.R.readerResult;


import cn.creekmoon.excel.util.exception.ExConsumer;

import java.time.LocalDateTime;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

/**
 * 读取结果
 */
public abstract class ReaderResult<R> {

    public LocalDateTime readStartTime;
    public LocalDateTime readSuccessTime;
    public LocalDateTime consumeSuccessTime;
    /*错误次数统计*/
    public AtomicInteger errorCount = new AtomicInteger(0);
    public StringBuilder errorReport = new StringBuilder();
    /*存在读取失败的数据*/
    public AtomicReference<Boolean> EXISTS_READ_FAIL = new AtomicReference<>(false);


    abstract public ReaderResult<R> consume(ExConsumer<R> consumer) throws Exception;
}
