package cn.jy.excelUtils;

/**
 * @author JY
 * @date 2022/1/11
 */
@FunctionalInterface
public interface ExConsumer<T>{

    void accept(T t) throws Exception;
}
