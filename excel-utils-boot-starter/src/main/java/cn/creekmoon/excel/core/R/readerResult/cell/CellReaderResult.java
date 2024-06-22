package cn.creekmoon.excel.core.R.readerResult.cell;

import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.util.exception.ExConsumer;

import java.time.LocalDateTime;

public class CellReaderResult<R> extends ReaderResult<R> {


    private R data = null;


    @Override
    public ReaderResult<R> consume(ExConsumer<R> consumer) throws Exception {
        if (getData() == null) {
            return this;
        }
        consumer.accept(getData());
        consumeSuccessTime = LocalDateTime.now();
        return this;
    }


    public R getData() {
        if (EXISTS_READ_FAIL.get()) {
            return null;
        }
        return data;
    }


    public void setData(R data) {
        this.data = data;
    }
}
