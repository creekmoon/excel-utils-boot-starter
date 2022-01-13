package cn.jy.excelUtils;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
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

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelImport<R> {
    public static final String RESULT_TITLE = "导入结果";
    private ExcelReader currentReader;
    /* key=title  value=执行器 */
    private Map<String, ExFunction> converts = new HashMap(32);
    private Map<String, BiConsumer> consumers = new HashMap(32);
    /* key=title */
    private Set<String> mustExistTitles = new HashSet<>(32);
    private Set<String> skipEmptyTitles = new HashSet<>(32);
    /* 所有原生的Excel行 */
    private List<Map<String, Object>> rows;
    /* key=对象  value=未转换的对象  如果转换失败不会在里面*/
    private Map<R, Map<String, Object>> object2Row = new LinkedHashMap<>(32);

    private Supplier<R> newObjectSupplier;
    public R currentObject;


    private ExcelImport() {
    }


    public static <T> ExcelImport<T> create(MultipartFile file, Supplier<T> supplier) throws IOException {
        ExcelImport<T> excelImport = new ExcelImport();
        excelImport.newObjectSupplier = supplier;
        excelImport.currentReader = ExcelUtil.getReader(file.getInputStream());
        /*格式化器 将一切结果都转为String*/
        excelImport.currentReader.setCellEditor((Cell cell, Object value) -> {
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
        });
        return excelImport;
    }

    public <T> ExcelImport<R> addConvert(String title, BiConsumer<R, T> setter, ExFunction<String, T> convert) throws IOException {
        converts.put(title, convert);
        consumers.put(title, setter);
        return this;
    }

    public  ExcelImport<R> addConvert(String title, BiConsumer<R, String> setter) throws IOException {
        addConvert(title, setter, x -> x);
        return this;
    }


    public <T> ExcelImport<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter) throws IOException {
        skipEmptyTitles.add(title);
        addConvertAndMustExist(title, setter, x -> x);
        return this;
    }

    public <T> ExcelImport<R> addConvertAndSkipEmpty(String title, BiConsumer<R, T> setter, ExFunction<String, T> convert) throws IOException {
        skipEmptyTitles.add(title);
        addConvertAndMustExist(title, setter, convert);
        return this;
    }


    public ExcelImport<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter) throws IOException {
        mustExistTitles.add(title);
        addConvertAndMustExist(title, setter, x -> x);
        return this;
    }

    public <T> ExcelImport<R> addConvertAndMustExist(String title, BiConsumer<R, T> setter, ExFunction<String, T> convert) throws IOException {
        mustExistTitles.add(title);
        this.addConvert(title, setter, convert);
        return this;
    }

    public List<R> readAsList(boolean returnEmptyIfFail) {
        /*是否存在检测不通过的情况*/
        boolean existsCheckFail = false;
        rows = currentReader.readAll();
        for (Map<String, Object> row : rows) {
            currentObject = newObjectSupplier.get();
            object2Row.put(currentObject, row);
            try {
                rowCheckMustExist(row);
                rowConvert(row);
                row.put(RESULT_TITLE, getCheckSuccessMsg());
            } catch (MyException e) {
                existsCheckFail = true;
                row.put(RESULT_TITLE, e.getMessage());
                object2Row.remove(currentObject);
            }
        }
        if (existsCheckFail && returnEmptyIfFail) {
            return Collections.EMPTY_LIST;
        }
        return new ArrayList<R>(object2Row.keySet());

    }

    /* 读取文件  忽略转换失败的行 */
    public void wtxSimpleReadAndSkipErrorRow(HttpServletResponse response,ExConsumer<R> insert) throws IOException {
        simpleRead(response, insert, false);
    }

    /* 读取  遇到失败则停止 */
    public void wtxSimpleRead(HttpServletResponse response,ExConsumer<R> insert) throws IOException {
        simpleRead(response, insert, true);
    }

    /* 读取 */
    private void simpleRead(HttpServletResponse response,ExConsumer<R> insert, Boolean returnEmptyIfFail) throws IOException {
        List<R> read = readAsList(returnEmptyIfFail);
        for (R row : read) {
            try {
                insert.accept(row);
                this.setResult(row, getInsertSuccessMsg());
            } catch (MyException e) {
                this.setResult(row, e.getMessage());
            } catch (Exception e) {
                log.error("导入失败!", e);
                this.setResult(row, getErrorMsg());
            }
        }
        this.response(response);
    }


    public void response(HttpServletResponse response) throws IOException {
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH-mm");
        ExcelExport export = ExcelExport.create("result[" + LocalDateTime.now().format(dateTimeFormatter)+"]");
        export.writeByMap(rows);
        export.send(response);
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
    private void rowConvert(Map<String, Object> row) throws MyException {

        /*执行convert*/
        for (Map.Entry<String, Object> entry : row.entrySet()) {
            String value = Optional.ofNullable(entry.getValue()).map(x -> (String) x).orElse("");
            if (skipEmptyTitles.contains(entry.getKey()) && StrUtil.isBlank(value)) {
                continue;
            }
            try {
                Object convertValue = converts.get(entry.getKey()).apply(value);
                consumers.get(entry.getKey()).accept(currentObject, convertValue);
            } catch (MyException e) {
                throw e;
            } catch (Exception e) {
                e.printStackTrace();
                log.error("导入文件发生错误!", e);
                throw new MyException(getErrorMsg());
            }
        }

    }

    /*校验必填项*/
    private void rowCheckMustExist(Map<String, Object> row) throws MyException {

        /*检查必填项*/
        for (String key : row.keySet()) {
            if (mustExistTitles.contains(key)) {
                Object str = row.get(key);
                if (str == null || StrUtil.isBlank((String) str)) {
                    throw new MyException(getNonNullMsg(key));
                }
            }
        }
    }

    /*必填项没有填写时的提示信息*/
    private static String getNonNullMsg(String keyName) {
        return "[" + keyName + "]" + "为必填项!";
    }


    /*校验成功的提示信息*/
    private static String getCheckSuccessMsg() {
        return "校验通过!";
    }

    /*导入失败的提示信息*/
    private static String getErrorMsg() {
        return "未知导入异常!请联系管理员!";
    }

    /*校验成功的提示信息*/
    private static String getInsertSuccessMsg() {
        return "导入成功!";
    }

}
