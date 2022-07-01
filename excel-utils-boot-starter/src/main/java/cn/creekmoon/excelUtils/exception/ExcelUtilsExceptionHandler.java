package cn.creekmoon.excelUtils.exception;

import org.springframework.core.Ordered;

/**
 * 异常处理器
 */
public interface ExcelUtilsExceptionHandler extends Ordered {

    /**
     * 自定义异常结果
     * @param unCatchException 未捕获的异常
     * @return   返回错误信息
     */
    String customExceptionMessage(Exception unCatchException);


    /**
     * 顺序 越低越先执行
     * @return
     */
    default int getOrder(){
        return 0;
    };
}
