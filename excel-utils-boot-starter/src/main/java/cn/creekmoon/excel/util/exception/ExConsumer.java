package cn.creekmoon.excel.util.exception;

/**
 * @author JY
 * @date 2022/1/11
 */
@FunctionalInterface
public interface ExConsumer<T> {

    void accept(T t) throws Exception;

    public static final ExConsumer EMPTY = o -> {
    };

    public static ExConsumer empty() {
        return EMPTY;
    }
}
