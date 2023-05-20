package cn.creekmoon.excelUtils.core;

/**
 * @author JY
 * @date 2023-5-20 17:42:39
 */
@FunctionalInterface
public interface ExBiConsumer<T, U> {


    void accept(T t, U u) throws Exception;
}
