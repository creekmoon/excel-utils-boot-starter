package cn.creekmoon.excelUtils.hutool589.log.dialect.tinylog;

import cn.creekmoon.excelUtils.hutool589.log.Log;
import cn.creekmoon.excelUtils.hutool589.log.LogFactory;

/**
 * <a href="http://www.tinylog.org/">TinyLog2</a> log.<br>
 *
 * @author Looly
 */
public class TinyLog2Factory extends LogFactory {

    /**
     * 构造
     */
    public TinyLog2Factory() {
        super("TinyLog");
        checkLogExist(org.tinylog.Logger.class);
    }

    @Override
    public Log createLog(String name) {
        return new TinyLog2(name);
    }

    @Override
    public Log createLog(Class<?> clazz) {
        return new TinyLog2(clazz);
    }

}
