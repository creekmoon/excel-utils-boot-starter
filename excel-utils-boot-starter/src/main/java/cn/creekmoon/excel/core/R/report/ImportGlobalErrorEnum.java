package cn.creekmoon.excel.core.R.report;

import lombok.Getter;

/**
 * 导入全局错误枚举
 */
@Getter
public enum ImportGlobalErrorEnum {

    TEMPLATE_MISMATCH("导入的模板有误,请检查您的文件!"),
    IMPORT_INTERRUPTED("导入过程发生异常，已终止读取");

    private final String message;

    ImportGlobalErrorEnum(String message) {
        this.message = message;
    }
}
