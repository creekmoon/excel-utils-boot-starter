package cn.creekmoon.excel.util.exception;

/**
 * 三参 + 可抛受检异常
 * @author JY
 * @date 2025-11-23
 */
@FunctionalInterface
public interface ExTriConsumer<A, B, C> {

    /**
     * 业务逻辑入口，允许抛出 Exception
     */
    void accept(A a, B b, C c) throws Exception;
}