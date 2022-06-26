package cn.jy.excelUtils.core;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.lang.UUID;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
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

import static cn.jy.excelUtils.core.ExcelConstants.CONVERT_SUCCESS_MSG;
import static cn.jy.excelUtils.core.ExcelConstants.IMPORT_SUCCESS_MSG;

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
        reader.setCellEditor((Cell cell, Object value) -> object2String(value));
        return reader;
    }


    /**
     * 全局默认数据转换器 将所有的数值都转换成String
     *
     * @param value
     * @return
     */
    private String object2String(Object value) {
        if (value == null) {
            return "";
        }
        if (value instanceof String) {
            return StrUtil.trim((String) value);
        }
        if (value instanceof Date) {
            return DateUtil.format((Date) value, "yyyy-MM-dd HH:mm:ss");
        }
        return String.valueOf(value);
    }

    /**
     * 初始化SAX读取器
     * 将所有的数据按行读取
     *
     * @throws IOException
     */
    Excel07SaxReader initSaxReader(AsyncTaskState asyncTaskState, ExConsumer<R> dataConsumer) {
        /*转换器队列*/
        ArrayList<ExFunction> convertsList = new ArrayList<>(converts.values());
        /*setter队列*/
        ArrayList<BiConsumer> setterList = new ArrayList<>(consumers.values());

        /*返回一个Sax读取器*/
        return new Excel07SaxReader((sheetIndex, rowIndex, rowList) -> {
            /*只读取第一个sheet*/
            if (sheetIndex != 0) {
                return;
            }
            /*跳过第一行*/
            if (rowIndex == 0) {
                return;
            }
            /*读取当前行的每一个属性 最终得到rowData*/
            R rowData = newObjectSupplier.get();
            for (int colIndex = 0; colIndex < convertsList.size(); colIndex++) {
                /*先把所有读到的数据转为String*/
                String value = object2String(rowList.get(colIndex));
                /*获取当前的转换器*/
                ExFunction converter = convertsList.get(colIndex);
                BiConsumer setter = setterList.get(colIndex);
                Object apply = null;
                try {
                    apply = converter.apply(value);
                    setter.accept(rowData, apply);
                } catch (Exception e) {
                    /*遇到异常*/
                    String exceptionMsg = GlobalExceptionManager.getExceptionMsg(e);
                    asyncTaskState.getErrorReport().put(rowIndex, exceptionMsg);
                }
            }
            /*消费这个rowData*/
            try {
                asyncTaskState.tryRowIndex++;
                dataConsumer.accept(rowData);
                asyncTaskState.successRowIndex++;

            } catch (Exception e) {
                String exceptionMsg = GlobalExceptionManager.getExceptionMsg(e);
                asyncTaskState.getErrorReport().put(rowIndex, exceptionMsg);
            }
            /*如果异常数量达到指定的值时 直接中断本次导入*/
            if (ASYNC_IMPORT_FAIL >= 0 && asyncTaskState.getErrorReport().size() >= ASYNC_IMPORT_FAIL) {
                throw new RuntimeException("异常数量过多，终止导入。请检查Excel文件内容是否正确");
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
        List<R> read = read(readStrategy);
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

    public List<R> read() {
        // 读取Excel中的对象, 如果存在任何一行转换失败,则返回空List
        return read(ReadStrategy.RETURN_EMPTY_LIST_IF_EXIST_FAIL);
    }

    @SneakyThrows
    public List<R> read(ReadStrategy readStrategy) {
        semaphore.acquire();
        try {
            /*初始化读取器*/
            ExcelReader memoryReader = initMemoryReader();
            /*是否存在读取不通过的情况*/
            boolean existsFail = false;
            rows = memoryReader.readAll();
            for (Map<String, Object> row : rows) {
                currentObject = newObjectSupplier.get();
                object2Row.put(currentObject, row);
                try {
                    rowCheckMustExist(row);
                    rowConvert(row);
                    row.put(RESULT_TITLE, CONVERT_SUCCESS_MSG);
                } catch (Exception e) {
                    existsFail = true;
                    row.put(RESULT_TITLE, GlobalExceptionManager.getExceptionMsg(e));
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


    /*行数据转换*/
    private void rowConvert(Map<String, Object> row) throws Exception {

        /*执行convert*/
        for (Map.Entry<String, Object> entry : row.entrySet()) {
            String value = Optional.ofNullable(entry.getValue()).map(x -> (String) x).orElse("");
            if (skipEmptyTitles.contains(entry.getKey()) && StrUtil.isBlank(value)) {
                continue;
            }
            Object convertValue = converts.get(entry.getKey()).apply(value);
            consumers.get(entry.getKey()).accept(currentObject, convertValue);
        }
    }

    /*校验必填项*/
    private void rowCheckMustExist(Map<String, Object> row) throws CheckedExcelException {

        /*检查必填项*/
        for (String key : row.keySet()) {
            if (mustExistTitles.contains(key)) {
                Object str = row.get(key);
                if (str == null || StrUtil.isBlank((String) str)) {
                    throw new CheckedExcelException(key + "为必填项!");
                }
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
