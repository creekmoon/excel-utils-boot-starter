package cn.creekmoon.excel.core.R.reader;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.util.exception.ExConsumer;
import lombok.Getter;

import java.io.IOException;
import java.util.function.Supplier;

public abstract class Reader<R> {

    public Reader(ExcelImport parent) {
        this.parent = parent;
    }

    //读取器持有其父类
    @Getter
    ExcelImport parent;

    public Supplier newObjectSupplier;

    /**
     * Sheet的rId（Excel内部关系ID，稳定可靠）
     */
    public String sheetRid;

    /**
     * Sheet的名称
     */
    public String sheetName;


}
