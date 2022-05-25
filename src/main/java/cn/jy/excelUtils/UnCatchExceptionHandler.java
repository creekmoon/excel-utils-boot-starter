package cn.jy.excelUtils;

/**
 * 未捕获异常处理器
 */
public interface UnCatchExceptionHandler {

    /**
     * 自定义异常结果
     * @param unCatchException 未捕获的异常
     * @return   返回错误信息
     */
    String customExceptionMessage(Exception unCatchException);
}
