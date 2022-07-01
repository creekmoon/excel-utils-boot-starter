package cn.creekmoon.excelUtils.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * ES连接常量
 */
@Component //定义配置类
@Data //提供get set方法
@ConfigurationProperties(prefix = "excel-utils") //yml配置中的路径
public class ExcelUtilsConfig {
    /**
     * 能并行执行多少个导入任务 防止内存溢出
     */
    private int importMaxParallel = 4;
    /**
     * 异步读写时,回调状态的刷新时间 单位毫秒
     */
    private int asyncRefreshMilliseconds = 1500;
    /**
     * 使用异步导入时,如果导入累计出现异常超过最大次数,则中断导入
     */
    private int asyncImportMaxFail = 100;
    /**
     * 临时文件的保留寿命 单位分钟
     */
    private int tempFileLifeMinutes = 5;
}
