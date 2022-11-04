package cn.creekmoon.excelUtils.hutool589.log.dialect.console;

import cn.creekmoon.excelUtils.hutool589.log.Log;
import cn.creekmoon.excelUtils.hutool589.log.LogFactory;

/**
 * 利用System.out.println()打印彩色日志
 *
 * @author hongda.li
 * @since 5.8.0
 */
public class ConsoleColorLogFactory extends LogFactory {

    public ConsoleColorLogFactory() {
        super("Hutool Console Color Logging");
    }

    @Override
    public Log createLog(String name) {
        return new ConsoleColorLog(name);
    }

    @Override
    public Log createLog(Class<?> clazz) {
        return new ConsoleColorLog(clazz);
    }
}
