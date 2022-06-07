package cn.jy.excelUtils.exception;

import lombok.extern.slf4j.Slf4j;
import org.springframework.context.annotation.Bean;
import org.springframework.core.Ordered;
import org.springframework.stereotype.Component;

@Slf4j
@Component
public class DefaultExcelUtilsExceptionHandler implements ExcelUtilsExceptionHandler {
    @Override
    public String customExceptionMessage(Exception unCatchException) {
        if (unCatchException instanceof CheckedExcelException) {
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
