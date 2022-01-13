package cn.jy.excelUtils;

/**
 * @author JY
 * @date 2022/1/11
 */
@FunctionalInterface
public interface ExFunction<T, R>{

    R apply(T t) throws Exception;
}
