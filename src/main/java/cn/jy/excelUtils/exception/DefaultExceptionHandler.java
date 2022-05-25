package cn.jy.excelUtils.exception;

import cn.jy.excelUtils.core.ExcelReadException;

public class DefaultExceptionHandler implements ExceptionHandler{
    @Override
    public String customExceptionMessage(Exception unCatchException) {
        if(unCatchException instanceof ExcelReadException){
            return unCatchException.getMessage();
        }
        return null;
    }

    @Override
    public Integer getLevel() {
        return Integer.MIN_VALUE;
    }
}
