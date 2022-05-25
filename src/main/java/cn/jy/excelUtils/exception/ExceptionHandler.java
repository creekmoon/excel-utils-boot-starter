package cn.jy.excelUtils.exception;

/**
 * 异常处理器
 */
public interface ExceptionHandler {

    /**
     * 自定义异常结果
     * @param unCatchException 未捕获的异常
     * @return   返回错误信息
     */
    String customExceptionMessage(Exception unCatchException);


    /**
     * 等级越高 优先级越高
     * @return
     */
    default Integer getLevel(){
        return 0;
    };
}
