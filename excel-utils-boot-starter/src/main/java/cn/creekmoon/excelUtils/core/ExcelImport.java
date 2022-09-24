package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.converter.MataConverter;
import cn.creekmoon.excelUtils.exception.CheckedExcelException;
import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;
import cn.hutool.core.text.StrFormatter;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
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
import java.util.function.Supplier;

import static cn.creekmoon.excelUtils.core.ExcelConstants.*;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelImport<R> {
    /**
     * 控制导入并发数量
     */
    public static Semaphore importSemaphore;
    /**
     * 导入结果的title
     */
    public static final String RESULT_TITLE = "导入结果";
    /**
     * 转换器策略
     */
    public ConvertStrategy convertStrategy;

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
        return create(file, supplier, ConvertStrategy.STOP_IF_CONVERT_ERROR);
    }


    public static <T> ExcelImport<T> create(MultipartFile file, Supplier<T> supplier, ConvertStrategy convertStrategy) {
        if (importSemaphore == null) {
            throw new RuntimeException("请使用@EnableExcelUtils进行初始化配置!");
        }
        ExcelImport<T> excelImport = new ExcelImport();
        excelImport.newObjectSupplier = supplier;
        excelImport.file = file;
        excelImport.convertStrategy = convertStrategy;
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
        addConvert(title, x -> x, setter);
        return this;
    }

    public <T> ExcelImport<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        skipEmptyTitles.add(title);
        addConvert(title, convert, setter);
        return this;
    }


    public ExcelImport<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter) {
        mustExistTitles.add(title);
        addConvert(title, x -> x, setter);
        return this;
    }

    public <T> ExcelImport<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        mustExistTitles.add(title);
        addConvert(title, convert, setter);
        return this;
    }

    public ExcelImport<R> read(ExConsumer<R> dataConsumer) {
        List<R> read = readAll();
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

    /**
     * 读取所有Excel到一个List
     *
     * @return
     */
    @SneakyThrows
    public List<R> readAll() {
        importSemaphore.acquire();
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
            if (convertStrategy == ConvertStrategy.STOP_IF_CONVERT_ERROR) {
                return existsFail ? Collections.EMPTY_LIST : new ArrayList<R>(object2Row.keySet());
            }
            /*如果读取策略为CONTINUE_ON_FAIL*/
            if (convertStrategy == ConvertStrategy.SKIP_FAIL_IF_CONVERT_ERROR) {
                return new ArrayList<R>(object2Row.keySet());
            }
            throw new RuntimeException("读取文档时发生错误,未定义读取策略ReadStrategy");
        } finally {
            importSemaphore.release();
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
        export.setColumnWidthDefault();
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
        /*最大转换次数*/
        int maxConvertCount = consumers.keySet().size();
        /*执行convert*/
        for (Map.Entry<String, Object> entry : row.entrySet()) {
            /*如果超过最大次数 不再进行读取*/
            if (maxConvertCount-- <= 0) {
                break;
            }
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
                log.warn("EXCEL导入数据转换失败！", e);
                throw new CheckedExcelException(StrFormatter.format(ExcelConstants.CONVERT_FAIL_MSG + GlobalExceptionManager.getExceptionMsg(e), entry.getKey()));
            }
        }
    }


    /**
     * 转换器策略
     */
    public enum ConvertStrategy {
        /*如果convert阶段失败, 跳过所有的行*/
        STOP_IF_CONVERT_ERROR,
        /*如果convert阶段失败, 跳过失败的行*/
        SKIP_FAIL_IF_CONVERT_ERROR;
    }

}
