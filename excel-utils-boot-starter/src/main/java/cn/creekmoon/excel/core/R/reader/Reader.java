package cn.creekmoon.excel.core.R.reader;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
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

    public ReaderResult readerResult;

    public Supplier newObjectSupplier;

    /**
     * Sheet的rId（Excel内部关系ID，稳定可靠）
     */
    public String sheetRid;

    /**
     * Sheet的名称
     */
    public String sheetName;

    /*启用模板一致性检查 为了防止模板导入错误*/
    public boolean TEMPLATE_CONSISTENCY_CHECK_ENABLE = true;

    /*标志位, 模板一致性检查已经失败 */
    public boolean TEMPLATE_CONSISTENCY_CHECK_FAILED = false;

    public abstract ReaderResult getReadResult();

    public abstract ReaderResult<R> read(ExConsumer<R> consumer) throws Exception;

    public abstract ReaderResult<R> read() throws InterruptedException, IOException;


}
