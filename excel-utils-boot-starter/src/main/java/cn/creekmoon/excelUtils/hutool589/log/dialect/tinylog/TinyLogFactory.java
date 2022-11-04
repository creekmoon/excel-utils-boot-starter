package cn.creekmoon.excelUtils.hutool589.log.dialect.tinylog;

import cn.creekmoon.excelUtils.hutool589.log.Log;
import cn.creekmoon.excelUtils.hutool589.log.LogFactory;

/**
 * <a href="http://www.tinylog.org/">TinyLog</a> log.<br>
 *
 * @author Looly
 */
public class TinyLogFactory extends LogFactory {

    /**
     * 构造
     */
    public TinyLogFactory() {
        super("TinyLog");
        checkLogExist(org.pmw.tinylog.Logger.class);
    }

    @Override
    public Log createLog(String name) {
        return new TinyLog(name);
    }

    @Override
    public Log createLog(Class<?> clazz) {
        return new TinyLog(clazz);
    }

}
