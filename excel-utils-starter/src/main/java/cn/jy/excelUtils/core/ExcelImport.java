package cn.jy.excelUtils.core;

import cn.hutool.core.lang.UUID;
import cn.hutool.core.text.StrFormatter;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import cn.jy.excelUtils.converter.MataConverter;
import cn.jy.excelUtils.exception.CheckedExcelException;
import cn.jy.excelUtils.exception.GlobalExceptionManager;
import cn.jy.excelUtils.threadPool.AsyncImportExecutor;
import cn.jy.excelUtils.threadPool.AsyncStateCallbackExecutor;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.Semaphore;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Supplier;

import static cn.jy.excelUtils.core.ExcelConstants.*;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelImport<R> {
    /**
     * 控制导入并发数量
     */
    public static Semaphore semaphore;
    /**
     * 导入结果的title
     */
    public static final String RESULT_TITLE = "导入结果";
    /**
     * 导入结果的title
     */
    public static int ASYNC_IMPORT_FAIL = -1;


    /* key=title  value=执行器 */
    private Map<String, ExFunction> converts = new LinkedHashMap(32);
    /* key=title value=消费者(通常是setter方法)*/
    private Map<String, BiConsumer> consumers = new LinkedHashMap(32);
    /* key=title */
    private Set<String> mustExistTitles = new HashSet<>(32);
    private Set<String> skipEmptyTitles = new HashSet<>(32);
    /* 所有原生的Excel行 */
    private List<Map<String, Object>> rows;
    /* key=对象  value=未转换的对象  如果转换失败不会在里面*/
    private Map<R, Map<String, Object>> object2Row = new LinkedHashMap<>(512);

    private Supplier<R> newObjectSupplier;
    /*当前导入的文件*/
    private MultipartFile file;
    public R currentObject;


    private ExcelImport() {
    }


    public static <T> ExcelImport<T> create(MultipartFile file, Supplier<T> supplier) {
        if (semaphore == null) {
            throw new RuntimeException("请使用@EnableExcelUtils进行初始化配置!");
        }
        ExcelImport<T> excelImport = new ExcelImport();
        excelImport.newObjectSupplier = supplier;
        excelImport.file = file;
        return excelImport;
    }


    /**
     * 初始化内存读取器
     * 将所有的数据读取至内存
     *
     * @throws IOException
     */
    private ExcelReader initMemoryReader() throws IOException {
        ExcelReader reader = ExcelUtil.getReader(file.getInputStream());
        /*格式化器 将一切结果都转为String*/
        reader.setCellEditor((Cell cell, Object value) -> MataConverter.parse(value));
        return reader;
    }


    /**
     * 初始化SAX读取器
     * 将所有的数据按行读取
     *
     * @throws IOException
     */
    Excel07SaxReader initSaxReader(AsyncTaskState asyncTaskState, ExConsumer<R> dataConsumer) {
        ExcelExport<Object> excelExport = ExcelExport.create("importResult");
        excelExport.taskId = asyncTaskState.taskId;
        /*转换器名称队列*/
        List<String> convertNamesList = new ArrayList<>(converts.keySet());

        /*返回一个Sax读取器*/
        return new Excel07SaxReader(new RowHandler() {
            @Override
            public void doAfterAllAnalysed() {
                /*sheet读取结束时*/
                excelExport.stopWrite();
                ExcelExport.cleanTempFile(excelExport.taskId);
            }

            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {
                /*只读取第一个sheet 并从第二行开始  因为第一行约定是标题*/
                if (sheetIndex != 0 || rowIndex == 0) {
                    return;
                }
                /*Excel解析原生的数据*/
                HashMap<String, Object> rawRowData = new LinkedHashMap<>();
                for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                    rawRowData.put(convertNamesList.get(colIndex), MataConverter.parse(rowList.get(colIndex)));
                }
                /*转换成业务对象*/
                try {
                    /*转换*/
                    asyncTaskState.tryRowCount++;
                    rowConvert(rawRowData);
                    rawRowData.put(RESULT_TITLE, CONVERT_SUCCESS_MSG);
                    /*消费*/
                    dataConsumer.accept(currentObject);
                    asyncTaskState.successRowCount++;
                    rawRowData.put(RESULT_TITLE, IMPORT_SUCCESS_MSG);
                } catch (Exception e) {
                    /*遇到异常 获取异常信息*/
                    String exceptionMsg = GlobalExceptionManager.getExceptionMsg(e);
                    /*写入异步状态 行号+1是因为Excel没有第0行*/
                    asyncTaskState.getErrorReport().put(rowIndex + 1, exceptionMsg);
                    /*写入导出Excel结果*/
                    rawRowData.put(RESULT_TITLE, exceptionMsg);
                } finally {
                    excelExport.writeByMap(Collections.singletonList(rawRowData));
                }
                /*如果异常数量达到指定的值时 直接中断本次导入*/
                if (ASYNC_IMPORT_FAIL >= 0 && asyncTaskState.getErrorReport().size() >= ASYNC_IMPORT_FAIL) {
                    excelExport.stopWrite();
                    ExcelExport.cleanTempFile(excelExport.taskId);
                    throw new RuntimeException("异常数量过多，终止导入。请检查Excel文件内容是否正确");
                }
            }
        });
    }

    public <T> ExcelImport<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        converts.put(title, convert);
        consumers.put(title, setter);
        return this;
    }

    public ExcelImport<R> addConvert(String title, BiConsumer<R, String> setter) {
        addConvert(title, x -> x, setter);
        return this;
    }


    public <T> ExcelImport<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter) {
        skipEmptyTitles.add(title);
        addConvertAndMustExist(title, x -> x, setter);
        return this;
    }

    public <T> ExcelImport<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) throws IOException {
        skipEmptyTitles.add(title);
        addConvertAndMustExist(title, convert, setter);
        return this;
    }


    public ExcelImport<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter) {
        mustExistTitles.add(title);
        addConvertAndMustExist(title, x -> x, setter);
        return this;
    }

    public <T> ExcelImport<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        mustExistTitles.add(title);
        this.addConvert(title, convert, setter);
        return this;
    }

    /**
     * 异步SAX方式读取，任务会提交到线程池中,注意将可能产生事务失效的情况
     */
    @SneakyThrows
    public AsyncTaskState readAsync(ExConsumer<R> dataConsumer, Consumer<AsyncTaskState> asyncCallback) {
        /*创建回调任务*/
        AsyncTaskState asyncTaskState = AsyncStateCallbackExecutor.createAsyncTaskState(UUID.fastUUID().toString(), asyncCallback);
        /*生成一个SAX读取器*/
        Excel07SaxReader saxReader = initSaxReader(asyncTaskState, dataConsumer);
        /*尝试拿锁*/
        semaphore.acquire();
        /*提交给异步线程池*/
        AsyncImportExecutor.run(() -> {
            try {
                /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
                saxReader.read(file.getInputStream(), -1);
            } catch (Exception e) {
                log.error("SaxReader读取Excel文件异常", e);
                e.printStackTrace();
            } finally {
                AsyncStateCallbackExecutor.completedAsyncTaskState(asyncTaskState.taskId);
                /*释放信号量*/
                semaphore.release();
            }
        });
        return asyncTaskState;
    }

    public ExcelImport<R> read(ExConsumer<R> dataConsumer) {
        return read(dataConsumer, ReadStrategy.RETURN_EMPTY_LIST_IF_EXIST_FAIL);
    }

    public ExcelImport<R> read(ExConsumer<R> dataConsumer, ReadStrategy readStrategy) {
        List<R> read = readAll(readStrategy);
        for (int i = 0; i < read.size(); i++) {
            try {
                dataConsumer.accept(read.get(i));
                setResult(read.get(i), IMPORT_SUCCESS_MSG);
            } catch (Exception e) {
                setResult(read.get(i), GlobalExceptionManager.getExceptionMsg(e));
            }
        }
        return this;
    }

    @SneakyThrows
    public List<R> readAll(ReadStrategy readStrategy) {
        semaphore.acquire();
        try {
            /*初始化读取器*/
            ExcelReader memoryReader = initMemoryReader();
            /*是否存在读取不通过的情况*/
            boolean existsFail = false;
            rows = memoryReader.readAll();
            for (Map<String, Object> row : rows) {
                try {
                    rowConvert(row);
                    /*转换成功*/
                    object2Row.put(currentObject, row);
                    row.put(RESULT_TITLE, CONVERT_SUCCESS_MSG);
                } catch (Exception e) {
                    existsFail = true;
                    row.put(RESULT_TITLE, GlobalExceptionManager.getExceptionMsg(e));
                    /*转换失败*/
                    object2Row.remove(currentObject);
                }
            }
            /*如果读取策略为RETURN_EMPTY_ON_FAIL*/
            if (readStrategy == ReadStrategy.RETURN_EMPTY_LIST_IF_EXIST_FAIL) {
                return existsFail ? Collections.EMPTY_LIST : new ArrayList<R>(object2Row.keySet());
            }
            /*如果读取策略为CONTINUE_ON_FAIL*/
            if (readStrategy == ReadStrategy.RETURN_SUCCESS_LIST_IF_EXIST_FAIL) {
                return new ArrayList<R>(object2Row.keySet());
            }
            throw new RuntimeException("读取文档时发生错误,未定义读取策略ReadStrategy");
        } finally {
            semaphore.release();
        }

    }


    /* 读取 */
    public void readAndResponse(ExConsumer<R> rowConsumer, HttpServletResponse response) throws IOException {
        this.read(rowConsumer);
        this.response(response);
    }

    public void response(HttpServletResponse response) throws IOException {
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH-mm");
        ExcelExport export = ExcelExport.create("result[" + LocalDateTime.now().format(dateTimeFormatter) + "]");
        export.writeByMap(rows);
        export.response(response);
    }

    /**
     * 设置读取的结果
     *
     * @param object 实例化的对象
     * @param msg    结果
     */
    public void setResult(R object, String msg) {
        Map<String, Object> row = object2Row.get(object);
        if (row != null) {
            row.put(RESULT_TITLE, msg);
        }
    }


    /**
     * 行转换
     *
     * @param row 实际上是Map<String, String>对象
     * @throws Exception
     */
    private void rowConvert(Map<String, Object> row) throws Exception {
        /*初始化空对象*/
        currentObject = newObjectSupplier.get();
        /*执行convert*/
        for (Map.Entry<String, Object> entry : row.entrySet()) {

            String value = Optional.ofNullable(entry.getValue()).map(x -> (String) x).orElse("");
            /*检查必填项/检查可填项*/
            if (StrUtil.isBlank(value)) {
                if (mustExistTitles.contains(entry.getKey())) {
                    throw new CheckedExcelException(StrFormatter.format(FIELD_LACK_MSG, entry.getKey()));
                }
                if (skipEmptyTitles.contains(entry.getKey())) {
                    continue;
                }
            }
            /*转换数据*/
            try {
                Object convertValue = converts.get(entry.getKey()).apply(value);
                consumers.get(entry.getKey()).accept(currentObject, convertValue);
            } catch (Exception e) {
                log.error("EXCEL导入数据转换失败！", e);
                throw new CheckedExcelException(StrFormatter.format(ExcelConstants.CONVERT_FAIL_MSG, entry.getKey()));
            }
        }
    }


    /**
     * 读取策略
     */
    public enum ReadStrategy {
        /*遇到失败时返回空行*/
        RETURN_EMPTY_LIST_IF_EXIST_FAIL,
        /*遇到失败时跳过异常的行 并返回成功的行*/
        RETURN_SUCCESS_LIST_IF_EXIST_FAIL;
    }

}
