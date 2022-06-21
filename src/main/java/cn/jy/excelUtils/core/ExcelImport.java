package cn.jy.excelUtils.core;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import cn.jy.excelUtils.exception.CheckedExcelException;
import cn.jy.excelUtils.exception.GlobalExceptionManager;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

import static cn.jy.excelUtils.core.ExcelConstants.*;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelImport<R> {
    public static final String RESULT_TITLE = "导入结果";
    /**
     * 默认读取器-基于内存
     */
    private ExcelReader currentReader;
    /**
     * Sax读取器-基于流
     */
    private Excel07SaxReader currentSaxReader;
    /* key=title  value=执行器 */
    private Map<String, ExFunction> converts = new LinkedHashMap(32);
    private Map<String, BiConsumer> consumers = new LinkedHashMap(32);
    /* key=title */
    private Set<String> mustExistTitles = new HashSet<>(32);
    private Set<String> skipEmptyTitles = new HashSet<>(32);
    /* 所有原生的Excel行 */
    private List<Map<String, Object>> rows;
    /* key=对象  value=未转换的对象  如果转换失败不会在里面*/
    private Map<R, Map<String, Object>> object2Row = new LinkedHashMap<>(32);

    private Supplier<R> newObjectSupplier;
    /*当前导入的文件*/
    private MultipartFile file;
    public R currentObject;


    private ExcelImport() {
    }


    public static <T> ExcelImport<T> create(MultipartFile file, Supplier<T> supplier) {
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
    private void initMemoryReader() throws IOException {
        if (this.currentReader != null) {
            return;
        }
        this.currentReader = ExcelUtil.getReader(file.getInputStream());
        /*格式化器 将一切结果都转为String*/
        this.currentReader.setCellEditor((Cell cell, Object value) -> object2String(value));
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
     * 初始化内存读取器
     * 将所有的数据读取至内存
     *
     * @throws IOException
     */
    private SaxReaderStatus initSaxReader(ExConsumer<R> dataConsumer) {

        SaxReaderStatus saxReaderStatus = new SaxReaderStatus();
        if (this.currentSaxReader != null) {
            return saxReaderStatus;
        }
        /*转换器队列 按下标而不是按K-V形式*/
        ArrayList<ExFunction> convertsList = new ArrayList<>(converts.values());
        /*转换器队列 按下标而不是按K-V形式*/
        ArrayList<BiConsumer> consumerList = new ArrayList<>(consumers.values());
        this.currentSaxReader = new Excel07SaxReader(new RowHandler() {
            @Override
            public void handle(int sheetIndex, int rowIndex, List<Object> rowList) {
                /*只读取第一个sheet*/
                if (sheetIndex != 0) {
                    return;
                }
                /*跳过第一行*/
                if (rowIndex == 0) {
                    return;
                }
                R rowData = newObjectSupplier.get();
                /*将当前行转为对象*/
                for (int colIndex = 0; colIndex < convertsList.size(); colIndex++) {
                    /*先把所有读到的数据转为String*/
                    String value = object2String(rowList.get(colIndex));
                    /*获取当前的转换器*/
                    ExFunction converter = convertsList.get(colIndex);
                    BiConsumer setter = consumerList.get(colIndex);

                    Object apply = null;
                    try {
                        apply = converter.apply(value);
                        setter.accept(rowData, apply);
                    } catch (Exception e) {
                        String exceptionMsg = GlobalExceptionManager.getExceptionMsg(e);
                        saxReaderStatus.getErrorReport().put(rowIndex, exceptionMsg);
                    }
                }
                /*消费这个对象 通常就是insert*/
                try {
                    dataConsumer.accept(rowData);
                    saxReaderStatus.successRowIndex++;
                } catch (Exception e) {
                    String exceptionMsg = GlobalExceptionManager.getExceptionMsg(e);
                    saxReaderStatus.getErrorReport().put(rowIndex, exceptionMsg);
                }
                /*如果异常数量达到100 直接中断本次导入*/
                if (saxReaderStatus.getErrorReport().size() >= 100) {
                    throw new RuntimeException("异常数量过多，终止导入。请检查Excel文件格式");
                }
                //Console.log("[{}] [{}] {}", sheetIndex, rowIndex, rowList);
            }
        });
        return saxReaderStatus;
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
     * SAX方式读取
     */
    public SaxReaderStatus saxRead(ExConsumer<R> dataConsumer) throws IOException {
        SaxReaderStatus saxReaderStatus = initSaxReader(dataConsumer);
        try {
            /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
            this.currentSaxReader.read(file.getInputStream(), -1);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return saxReaderStatus;
    }

    public ExcelImport<R> read(ExConsumer<R> dataConsumer) throws IOException {
        return read(dataConsumer, ReadStrategy.RETURN_EMPTY_LIST_IF_EXIST_FAIL);
    }

    public ExcelImport<R> read(ExConsumer<R> dataConsumer, ReadStrategy readStrategy) throws IOException {
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

    public List<R> read() throws IOException {
        // 读取Excel中的对象, 如果存在任何一行转换失败,则返回空List
        return read(ReadStrategy.RETURN_EMPTY_LIST_IF_EXIST_FAIL);
    }

    public List<R> read(ReadStrategy readStrategy) throws IOException {
        /*初始化读取器*/
        initMemoryReader();
        /*是否存在读取不通过的情况*/
        boolean existsFail = false;
        rows = currentReader.readAll();
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
    }


    /* 读取 */
    public void readAndResponse(HttpServletResponse response, ExConsumer<R> rowConsumer) throws IOException {
        List<R> read = read();
        for (R row : read) {
            try {
                rowConsumer.accept(row);
                this.setResult(row, IMPORT_SUCCESS_MSG);
            } catch (Exception e) {
                log.error(ERROR_MSG, e);
                this.setResult(row, GlobalExceptionManager.getExceptionMsg(e));
            }
        }
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
