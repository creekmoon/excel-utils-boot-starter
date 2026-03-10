package cn.creekmoon.excel.core.R.report;

import lombok.Data;

/**
 * Excel导入配置选项（唯一配置入口，避免散落开关）
 */
@Data
public class ImportOptions {

    /**
     * 模板错误时是否立即停止（fail-fast）
     * true：标题行校验不通过立即终止，不扫描数据行
     * false：继续读取（旧版兼容行为）
     * 默认：false
     */
    private boolean failFastOnTemplateMismatch = false;

    /**
     * 最大错误行数阈值（达到该值立即停止读取）
     * null：不限制
     * 默认：null
     */
    private Integer maxErrorRows = null;

    /**
     * 是否收集行级错误明细到 ImportReport.rowErrors
     * true：收集错误明细（可用于JSON返回/调试）
     * false：仅统计错误数量
     * 默认：true
     */
    private boolean collectRowErrors = true;

    /**
     * 是否收集成功数据到 ImportReport.data
     * true：收集成功对象（便于API返回/业务处理）
     * false：不收集（节省内存，适合只需response的场景）
     * 默认：false
     */
    private boolean collectSuccessData = false;

    /**
     * 成功数据收集上限
     * null：不限制
     * 仅当 collectSuccessData=true 时生效
     * 默认：null
     */
    private Integer maxSuccessData = null;


    /**
     * 创建默认配置
     */
    public static ImportOptions defaults() {
        return new ImportOptions();
    }

    /**
     * 创建API场景预设（fail-fast + 只收集错误）
     */
    public static ImportOptions forApi() {
        ImportOptions options = new ImportOptions();
        options.setFailFastOnTemplateMismatch(true);
        options.setMaxErrorRows(100); // 避免大量错误
        options.setCollectSuccessData(false);
        return options;
    }

    /**
     * 创建Excel回写场景预设（读完整表 + 收集所有错误）
     */
    public static ImportOptions forExcelResponse() {
        ImportOptions options = new ImportOptions();
        options.setFailFastOnTemplateMismatch(false);
        options.setCollectRowErrors(true);
        options.setCollectSuccessData(false);
        return options;
    }
}
