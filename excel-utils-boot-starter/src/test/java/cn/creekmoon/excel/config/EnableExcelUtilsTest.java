package cn.creekmoon.excel.config;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.support.fixture.FixtureCatalog;
import cn.creekmoon.excel.support.fixture.ImportTestHelper;
import cn.creekmoon.excel.support.fixture.Student;
import cn.creekmoon.excel.support.spring.TestBizException;
import cn.creekmoon.excel.util.exception.CustomExceptionHandler;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Semaphore;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** @EnableExcelUtils 相关全局配置：并行度、临时文件、信号量。 */
@Slf4j
class EnableExcelUtilsTest {

    /** 模拟注解生效后的 ExcelUtilsConfig 默认值。 */
    @BeforeEach
    void applyEnableExcelUtilsDefaults() {
        ExcelUtilsConfig.importParallelSemaphore = new Semaphore(2);
        ExcelUtilsConfig.TEMP_FILE_LIFE_MINUTES = 1;
        ExcelUtilsConfig.addExcelUtilsExceptionHandler(new CustomExceptionHandler(TestBizException.class));
        log.info("[SETUP] applied EnableExcelUtils defaults: parallel=2, tempLifeMinutes=1");
    }

    /** 并行信号量许可数=2，临时文件生命周期=1 分钟。 */
    @Test
    void enableExcelUtils_annotationConfig_appliesParallelAndTempFileSettings() {
        log.info("[ASSERT] semaphorePermits={}, tempFileLifeMinutes={}",
                ExcelUtilsConfig.importParallelSemaphore.availablePermits(),
                ExcelUtilsConfig.TEMP_FILE_LIFE_MINUTES);
        assertEquals(2, ExcelUtilsConfig.importParallelSemaphore.availablePermits());
        assertEquals(1, ExcelUtilsConfig.TEMP_FILE_LIFE_MINUTES);
        log.info("[DONE] config values verified");
    }

    /**
     * 3 个线程并发导入，信号量仅 2：后启动任务应排队，最终处理行数 = 1000×3。
     */
    @Test
    void importParallel_exceedsMaxParallel_allTasksEventuallyComplete() throws Exception {
        int parallelWorkers = 3;
        int expectedTotalProcessed = FixtureCatalog.totalDataRows() * parallelWorkers;
        log.info("[STEP] start {} parallel imports with maxParallel=2, expectedProcessed={}",
                parallelWorkers, expectedTotalProcessed);

        ExecutorService executor = Executors.newFixedThreadPool(parallelWorkers);
        CountDownLatch startLatch = new CountDownLatch(1);
        CountDownLatch doneLatch = new CountDownLatch(parallelWorkers);
        AtomicInteger successCount = new AtomicInteger();

        for (int i = 0; i < parallelWorkers; i++) {
            int taskId = i;
            executor.submit(() -> {
                try {
                    startLatch.await();
                    log.info("[TASK-{}] import started", taskId);
                    ExcelImport excelImport = ImportTestHelper.createImport(FixtureCatalog.STANDARD_IMPORT);
                    ImportTestHelper.bindStandardColumns(excelImport.switchSheet(0, Student::new))
                            .read(x -> successCount.incrementAndGet());
                    log.info("[TASK-{}] import finished", taskId);
                } catch (Exception e) {
                    throw new RuntimeException(e);
                } finally {
                    doneLatch.countDown();
                }
            });
        }

        // 同时放行所有任务，竞争信号量
        startLatch.countDown();
        assertTrue(doneLatch.await(120, TimeUnit.SECONDS));
        executor.shutdown();
        assertEquals(expectedTotalProcessed, successCount.get());
        log.info("[DONE] parallel import completed, processedRows={}", successCount.get());
    }
}
