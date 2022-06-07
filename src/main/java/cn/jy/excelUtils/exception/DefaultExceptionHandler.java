package cn.jy.excelUtils.exception;

import cn.jy.excelUtils.core.ExcelConstants;
import cn.jy.excelUtils.core.ExcelReadException;
import lombok.extern.slf4j.Slf4j;

@Slf4j
public class DefaultExceptionHandler implements ExceptionHandler {
    @Override
    public String customExceptionMessage(Exception unCatchException) {
        if (unCatchException instanceof ExcelReadException) {
            return unCatchException.getMessage();
        }
        //返回null说明不进行处理 将委托给其他处理器进行处理
        return null;
    }

    @Override
    public Integer getLevel() {
        return Integer.MIN_VALUE;
    }
}
