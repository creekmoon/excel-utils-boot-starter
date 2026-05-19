package cn.creekmoon.excel.util.exception;

import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;

import static org.junit.jupiter.api.Assertions.*;

/** CustomExceptionHandler：按类名匹配，不依赖 instanceof（兼容多 ClassLoader）。 */
@Slf4j
class CustomExceptionHandlerTest {

    /** 同名类不同 ClassLoader：isInstance 为 false，但 handler 仍能识别并取消息。 */
    @Test
    void customExceptionHandler_sameClassNameAcrossClassLoaders_matchesAndExtractsMessage() throws Exception {
        log.info("[STEP] create handler for CrossLoaderBizException");
        CustomExceptionHandler handler = new CustomExceptionHandler(CrossLoaderBizException.class);
        Exception crossLoaderException = createCrossLoaderBizException("跨类加载器业务异常");

        log.info("[RESULT] isInstance={}, handlerMatch={}, message={}",
                CrossLoaderBizException.class.isInstance(crossLoaderException),
                handler.isCustomException(crossLoaderException),
                handler.customExceptionMessage(crossLoaderException));
        assertFalse(CrossLoaderBizException.class.isInstance(crossLoaderException));
        assertTrue(handler.isCustomException(crossLoaderException));
        assertEquals("跨类加载器业务异常", handler.customExceptionMessage(crossLoaderException));
        log.info("[DONE] cross classloader match passed");
    }

    /** 只匹配异常本身类型，不沿 cause 链向下查找。 */
    @Test
    void customExceptionHandler_wrappedCause_doesNotMatchOuterException() {
        log.info("[STEP] wrap biz exception inside RuntimeException");
        CustomExceptionHandler handler = new CustomExceptionHandler(CrossLoaderBizException.class);
        RuntimeException wrappedException = new RuntimeException(new CrossLoaderBizException("被包装的业务异常"));

        log.info("[RESULT] outerMatch={}, message={}",
                handler.isCustomException(wrappedException),
                handler.customExceptionMessage(wrappedException));
        assertFalse(handler.isCustomException(wrappedException));
        assertNull(handler.customExceptionMessage(wrappedException));
        log.info("[DONE] wrapped exception not matched as expected");
    }

    private Exception createCrossLoaderBizException(String message) throws Exception {
        /*加载同名但不同 ClassLoader 的异常类*/
        String className = CrossLoaderBizException.class.getName();
        CrossLoaderClassLoader classLoader = new CrossLoaderClassLoader(className);
        Class<?> reloadedClass = classLoader.loadClass(className);

        /*反射创建异常实例*/
        try {
            var constructor = reloadedClass.getDeclaredConstructor(String.class);
            constructor.setAccessible(true);
            Object instance = constructor.newInstance(message);
            return (Exception) instance;
        } catch (InvocationTargetException e) {
            throw new RuntimeException("创建跨类加载器异常失败", e);
        }
    }

    /** 测试用业务异常（与生产侧「同名不同加载器」场景对齐）。 */
    static class CrossLoaderBizException extends Exception {
        public CrossLoaderBizException(String message) {
            super(message);
        }
    }

    /** 仅对目标类走 defineClass，其余委托父加载器。 */
    static class CrossLoaderClassLoader extends ClassLoader {

        private final String targetClassName;

        CrossLoaderClassLoader(String targetClassName) {
            super(CustomExceptionHandlerTest.class.getClassLoader());
            this.targetClassName = targetClassName;
        }

        @Override
        protected Class<?> loadClass(String name, boolean resolve) throws ClassNotFoundException {
            if (!targetClassName.equals(name)) {
                return super.loadClass(name, resolve);
            }
            Class<?> loadedClass = findLoadedClass(name);
            if (loadedClass == null) {
                byte[] classBytes = readTargetClassBytes(name);
                loadedClass = defineClass(name, classBytes, 0, classBytes.length);
            }
            if (resolve) {
                resolveClass(loadedClass);
            }
            return loadedClass;
        }

        private byte[] readTargetClassBytes(String className) throws ClassNotFoundException {
            String resourcePath = className.replace('.', '/') + ".class";
            try (InputStream inputStream = CustomExceptionHandlerTest.class.getClassLoader().getResourceAsStream(resourcePath)) {
                if (inputStream == null) {
                    throw new ClassNotFoundException("找不到类文件: " + className);
                }
                return inputStream.readAllBytes();
            } catch (IOException e) {
                throw new ClassNotFoundException("读取类文件失败: " + className, e);
            }
        }
    }
}
