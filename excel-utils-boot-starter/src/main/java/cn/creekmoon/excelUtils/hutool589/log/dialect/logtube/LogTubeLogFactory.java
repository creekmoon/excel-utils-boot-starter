package cn.creekmoon.excelUtils.hutool589.log.dialect.logtube;

import cn.creekmoon.excelUtils.hutool589.log.Log;
import cn.creekmoon.excelUtils.hutool589.log.LogFactory;

/**
 * <a href="https://github.com/logtube/logtube-java">LogTube</a> log. 封装<br>
 *
 * @author Looly
 * @since 5.6.6
 */
public class LogTubeLogFactory extends LogFactory {

    public LogTubeLogFactory() {
        super("LogTube");
        checkLogExist(io.github.logtube.Logtube.class);
    }

    @Override
    public Log createLog(String name) {
        return new LogTubeLog(name);
    }

    @Override
    public Log createLog(Class<?> clazz) {
        return new LogTubeLog(clazz);
    }

}
