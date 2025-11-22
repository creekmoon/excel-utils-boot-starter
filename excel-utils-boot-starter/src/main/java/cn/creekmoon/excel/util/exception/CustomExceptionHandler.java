package cn.creekmoon.excel.util.exception;

import lombok.SneakyThrows;
import org.springframework.core.Ordered;

/**
 * 业务异常处理器
 *
 *
 */
public class CustomExceptionHandler implements Ordered {

    /*自定义异常*/
    private Class customExceptionClass = null;

    @SneakyThrows
    public CustomExceptionHandler(String customExceptionClassName) {
        this.customExceptionClass = Class.forName(customExceptionClassName);
    }

    public CustomExceptionHandler(Class customExceptionClass) {
        this.customExceptionClass = customExceptionClass;
    }

    public boolean isCustomException(Exception catchException) {
        return customExceptionClass.isInstance(catchException);
    }

    public String customExceptionMessage(Exception unCatchException) {
        if (customExceptionClass.isInstance(unCatchException)) {
            return unCatchException.getMessage();
        }
        //返回null说明不进行处理 将委托给其他处理器进行处理
        return null;
    }


    @Override
    public int getOrder() {
        return Ordered.LOWEST_PRECEDENCE;
    }
}