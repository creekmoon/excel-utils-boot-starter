package cn.jy.excelUtils;

/**
 * 未捕获异常处理器
 */
public interface UnCatchExceptionHandler {

    /**
     * 获取异常信息内容
     * @param unCatchException 未捕获的异常
     * @return   返回错误信息
     */
    String exception2Message(Exception unCatchException);
}
