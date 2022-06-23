package cn.jy.excelUtils.config;

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
     * 使用SAXReader时, 异步返回的状态刷新频率 单位毫秒
     */
    private int saxReadRefreshMilliseconds = 1500;
    /**
     * 使用SAXReader时,如果导入累计出现异常超过最大次数,则中断导入
     */
    private int saxReadMaxFail = 100;
}
