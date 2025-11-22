package cn.creekmoon.excel.core.R.readerResult.title;

import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.exception.ExBiConsumer;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.GlobalExceptionMsgManager;
import cn.hutool.core.map.BiMap;
import cn.hutool.core.text.StrFormatter;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;


/**
 * 读取结果
 */
@Slf4j
public class TitleReaderResult<R> extends ReaderResult<R> {


    /*K=行下标 V=数据*/
    public BiMap<Integer, R> rowIndex2dataBiMap = new BiMap<>(new LinkedHashMap<>());

    /*行结果集合*/
    public LinkedHashMap<Integer, String> rowIndex2msg = new LinkedHashMap<>();

    public List<R> getAll() {
        if (EXISTS_READ_FAIL.get()) {
            // 如果转化阶段就存在失败数据, 意味着数据不完整,应该返回空
            return new ArrayList<>();
        }
        if (rowIndex2dataBiMap.isEmpty() && !rowIndex2msg.isEmpty()) {
            // 数据缓存已关闭，返回空列表
            log.warn("[Excel读取警告] 数据缓存已关闭(enableDataMemoryCache=false)，getAll()返回空列表。如需获取数据，请使用read(consumer)进行流式消费。");
        }
        return new ArrayList<>(rowIndex2dataBiMap.values());
    }


    public TitleReaderResult<R> consume(ExConsumer<R> dataConsumer) {
        return this.consume((index, data) -> dataConsumer.accept(data));
    }

    public TitleReaderResult<R> consume(ExBiConsumer<Integer, R> rowIndexAndDataConsumer) {
        if (rowIndex2dataBiMap.isEmpty() && !rowIndex2msg.isEmpty()) {
            // 数据缓存已关闭，无法批量消费
            log.warn("[Excel读取警告] 数据缓存已关闭(enableDataMemoryCache=false)，consume()方法无数据可消费。如需消费数据，请使用read(consumer)进行流式消费。");
            return this;
        }
        rowIndex2dataBiMap.forEach((rowIndex, data) -> {
            try {
                rowIndexAndDataConsumer.accept(rowIndex, data);
                rowIndex2msg.put(rowIndex, ExcelConstants.IMPORT_SUCCESS_MSG);
            } catch (Exception e) {
                errorCount.incrementAndGet();
                String exceptionMsg = GlobalExceptionMsgManager.getExceptionMsg(e);
                errorReport.append(StrFormatter.format("第[{}]行发生错误[{}]", (int) rowIndex + 1, exceptionMsg));
                rowIndex2msg.put(rowIndex, exceptionMsg);
            }
        });
        return this;
    }


    public TitleReaderResult<R> setResultMsg(R data, String msg) {
        Integer i = getDataIndexOrNull(data);
        if (i == null) {
            if (rowIndex2dataBiMap.isEmpty() && !rowIndex2msg.isEmpty()) {
                log.warn("[Excel读取警告] 数据缓存已关闭(enableDataMemoryCache=false)，无法通过数据对象反查行号设置消息。请使用setResultMsg(Integer rowIndex, String msg)方法。");
            }
            return this;
        }
        rowIndex2msg.put(i, msg);
        return this;
    }


    private  Integer getDataIndexOrNull(R data) {

        Integer i = rowIndex2dataBiMap.getKey(data);
        if (i == null) {
            log.error("[Excel读取异常]对象不在读取结果中! [{}]", data);
            return null;
        }
        return i;
    }


    public String getResultMsg(R data) {
        Integer i = getDataIndexOrNull(data);
        if (i == null) return null;
        return rowIndex2msg.get(i);
    }


    public String getResultMsg(Integer rowIndex) {
        return rowIndex2msg.get(rowIndex);
    }

    public TitleReaderResult<R> setResultMsg(Integer rowIndex, String msg) {
        if (rowIndex2msg.containsKey(rowIndex)) {
            rowIndex2msg.put(rowIndex, msg);
        }
        return this;
    }



    public AtomicInteger getErrorCount() {
        return errorCount;
    }

    public Integer getDataLatestRowIndex() {
        return this.rowIndex2msg.keySet().stream().max(Integer::compareTo).get();
    }

    public Integer getDataFirstRowIndex() {
        return this.rowIndex2msg.keySet().stream().min(Integer::compareTo).get();
    }
}
