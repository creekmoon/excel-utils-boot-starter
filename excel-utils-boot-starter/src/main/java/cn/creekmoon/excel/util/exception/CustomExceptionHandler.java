package cn.creekmoon.excel.util.exception;

import org.springframework.core.Ordered;

/**
 * 业务异常处理器
 *
 *
 */
public class CustomExceptionHandler implements Ordered {

    /*自定义异常类名*/
    private final String customExceptionClassName;

    public CustomExceptionHandler(String customExceptionClassName) {
        this.customExceptionClassName = customExceptionClassName;
    }

    public CustomExceptionHandler(Class<?> customExceptionClass) {
        this.customExceptionClassName = customExceptionClass.getName();
    }

    public boolean isCustomException(Exception catchException) {
        return isCurrentExceptionClassNameMatched(catchException);
    }

    public String customExceptionMessage(Exception unCatchException) {
        if (isCurrentExceptionClassNameMatched(unCatchException)) {
            return unCatchException.getMessage();
        }
        /*返回 null 说明不进行处理，将委托给其他处理器*/
        return null;
    }

    private boolean isCurrentExceptionClassNameMatched(Exception exception) {
        /*fast-fail*/
        if (exception == null) {
            return false;
        }

        /*只检查当前异常对象本身及其父类名，不下钻 cause 链*/
        Class<?> currentClass = exception.getClass();
        while (currentClass != null) {
            if (customExceptionClassName.equals(currentClass.getName())) {
                return true;
            }
            currentClass = currentClass.getSuperclass();
        }
        return false;
    }


    @Override
    public int getOrder() {
        return Ordered.LOWEST_PRECEDENCE;
    }
}