package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.config.ExcelUtilsConfig;
import cn.creekmoon.excelUtils.converter.MataConverter;
import cn.creekmoon.excelUtils.exception.CheckedExcelException;
import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;
import cn.creekmoon.excelUtils.hutool589.core.io.FileUtil;
import cn.creekmoon.excelUtils.hutool589.core.text.StrFormatter;
import cn.creekmoon.excelUtils.hutool589.core.text.csv.CsvReader;
import cn.creekmoon.excelUtils.hutool589.core.text.csv.CsvRow;
import cn.creekmoon.excelUtils.hutool589.core.util.StrUtil;
import cn.creekmoon.excelUtils.hutool589.poi.excel.ExcelReader;
import cn.creekmoon.excelUtils.hutool589.poi.excel.ExcelUtil;
import cn.creekmoon.excelUtils.hutool589.poi.excel.sax.Excel07SaxReader;
import cn.creekmoon.excelUtils.hutool589.poi.excel.sax.handler.RowHandler;
import cn.creekmoon.excelUtils.threadPool.CleanTempFilesExecutor;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStreamReader;
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
     * 标题行号 这里是0,意味着第一行是标题, 第二行才是数据
     */
    public int titleRowIndex = 0;

    /* key=title  value=执行器 */
    private LinkedHashMap<String, ExFunction> title2converts = new LinkedHashMap(32);
    /* key=title value=消费者(通常是setter方法)*/
    private LinkedHashMap<String, BiConsumer> title2consumers = new LinkedHashMap(32);
    /* key=title */
    private Set<String> mustExistTitles = new HashSet<>(32);
    private Set<String> skipEmptyTitles = new HashSet<>(32);
    private Supplier<R> newObjectSupplier;
    /*当前导入的文件*/
    private MultipartFile file;
    public R currentObject;
    /**
     * 导入结果
     */
    private ExcelExport excelExport;

    /*启用EXCEL标题模板一致性检查 为了防止标题重复 这里默认为true */
    public boolean ENABLE_TITLE_CHECK = true;
    /*如果标题检查出现失败, 则说明所有都失败, 这是一个性能优化. */
    private boolean FORCE_TITLE_CHECK_FAIL = false;
    //============================内存读取模式对象==========================
    /* 所有原生的Excel行 */
    private List<Map<String, Object>> rows;
    /* key=对象  value=未转换的对象  如果转换失败不会在里面*/
    private Map<R, Map<String, Object>> object2Row = new LinkedHashMap<>(512);
    //============================内存读取模式对象==========================

    private ExcelImport() {
    }


    public static void init() {
        if (importSemaphore == null) {
            importSemaphore = new Semaphore(ExcelUtilsConfig.IMPORT_MAX_PARALLEL);
        }
    }

    public static <T> ExcelImport<T> create(MultipartFile file, Supplier<T> supplier) {
        return create(file, supplier, ConvertStrategy.SKIP_ALL_IF_FAIL);
    }


    public static <T> ExcelImport<T> create(MultipartFile file, Supplier<T> supplier, ConvertStrategy convertStrategy) {
        if (importSemaphore == null) {
            throw new RuntimeException("请使用@EnableExcelUtils进行初始化配置!");
        }
        ExcelImport<T> excelImport = new ExcelImport();
        excelImport.newObjectSupplier = supplier;
        excelImport.file = file;

        /*初始化一个导入结果*/
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH-mm");
        excelImport.excelExport = ExcelExport.create("import_result[" + LocalDateTime.now().format(dateTimeFormatter) + "]");
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
        title2converts.put(title, convert);
        title2consumers.put(title, setter);
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


    @SneakyThrows
    public ExcelImport<R> read(ExConsumer<R> dataConsumer) {
        //旧版的读取 使用内存读取模式
//        List<R> read = readAll();
//        for (int i = 0; i < read.size(); i++) {
//            try {
//                dataConsumer.accept(read.get(i));
//                setResult(read.get(i), IMPORT_SUCCESS_MSG);
//            } catch (Exception e) {
//                setResult(read.get(i), GlobalExceptionManager.getExceptionMsg(e));
//            }
//        }

        /*尝试拿锁*/
        importSemaphore.acquire();
        try {
            this.csvSupport();
            //新版读取 使用SAX读取模式
            Excel07SaxReader excel07SaxReader = initSaxReader(dataConsumer);
            /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
            excel07SaxReader.read(file.getInputStream(), -1);
        } catch (Exception e) {
            log.error("SaxReader读取Excel文件异常", e);
            e.printStackTrace();
            ExcelExport.cleanTempFile(excelExport.stopWrite());
        } finally {
            /*释放信号量*/
            importSemaphore.release();
        }


        return this;
    }


    /**
     * 支持csv类型的文件  本质是内部将csv转为xlsx
     *
     * @return
     */
    @SneakyThrows
    public ExcelImport<R> csvSupport() {

        if (StrUtil.isBlank(this.file.getOriginalFilename()) || !this.file.getOriginalFilename().toLowerCase().contains(".csv")) {
            /*如果不是csv文件 跳过这个方法*/
            return this;
        }


        /*获取CSV文件并尝试读取*/
        ExcelExport<Object> newExcel = ExcelExport.create("csv2xlsx");
        try {
            CsvReader read = new CsvReader(new InputStreamReader(file.getInputStream()), null);
            Iterator<CsvRow> rowIterator = read.iterator();
            while (rowIterator.hasNext()) {
                newExcel.getBigExcelWriter().writeRow(rowIterator.next().getRawList());
            }
        } catch (Exception e) {
            log.error("csv转换异常!", e);
            e.printStackTrace();
            this.file = null;
        } finally {
            CleanTempFilesExecutor.cleanTempFileDelay(newExcel.stopWrite(), 10);
        }

        /*将新的xlsx文件替换为当前的文件*/
        this.file = new MockMultipartFile("csv2xlsx.xlsx", FileUtil.getInputStream(PathFinder.getAbsoluteFilePath(newExcel.taskId)));
        return this;
    }

    /**
     * 读取Excel内容到一个List (内存模式)
     *
     * @return
     */
    @SneakyThrows
    public List<R> readAll(ConvertStrategy convertStrategy) {
        importSemaphore.acquire();
        try {
            this.csvSupport();
            /*初始化读取器*/
            ExcelReader memoryReader = initMemoryReader();
            /*是否存在读取不通过的情况*/
            boolean existsFail = false;
            rows = memoryReader.read(titleRowIndex, titleRowIndex + 1, Integer.MAX_VALUE);
            for (Map<String, Object> row : rows) {
                try {
                    rowConvert(row);
                    row.put(RESULT_TITLE, CONVERT_SUCCESS_MSG);
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
            /*如果读取策略为SKIP_ALL_IF_FAIL*/
            if (convertStrategy == ConvertStrategy.SKIP_ALL_IF_FAIL) {
                return existsFail ? Collections.EMPTY_LIST : new ArrayList<R>(object2Row.keySet());
            }
            /*如果读取策略为SKIP_ROW_IF_FAIL*/
            if (convertStrategy == ConvertStrategy.SKIP_ROW_IF_FAIL) {
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
        if (rows != null && !rows.isEmpty()) {
            excelExport.writeByMap(rows);
        }
        excelExport.setColumnWidthDefault();
        excelExport.response(response);
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
        /*进行模板一致性检查*/
        if (ENABLE_TITLE_CHECK) {
            if (FORCE_TITLE_CHECK_FAIL || !titleConsistencyCheck(title2converts.keySet(), row.keySet())) {
                FORCE_TITLE_CHECK_FAIL = true;
                throw new CheckedExcelException(TITLE_CHECK_ERROR);
            }
        }
        ENABLE_TITLE_CHECK = false;

        /*初始化空对象*/
        currentObject = newObjectSupplier.get();
        /*最大转换次数*/
        int maxConvertCount = title2consumers.keySet().size();
        /*执行convert*/
        for (Map.Entry<String, Object> entry : row.entrySet()) {
            /*如果包含不支持的标题,  或者已经超过最大次数则不再进行读取*/
            if (!title2consumers.containsKey(entry.getKey()) || maxConvertCount-- <= 0) {
                continue;
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
                Object convertValue = title2converts.get(entry.getKey()).apply(value);
                title2consumers.get(entry.getKey()).accept(currentObject, convertValue);
            } catch (Exception e) {
                log.warn("EXCEL导入数据转换失败！", e);
                throw new CheckedExcelException(StrFormatter.format(ExcelConstants.CONVERT_FAIL_MSG + GlobalExceptionManager.getExceptionMsg(e), entry.getKey()));
            }
        }
    }


    /**
     * 标题一致性检查
     *
     * @param targetTitles 我们声明的要拿取的标题
     * @param sourceTitles 传过来的excel文件标题
     * @return
     */
    private Boolean titleConsistencyCheck(Set<String> targetTitles, Set<String> sourceTitles) {
        if (targetTitles.size() > sourceTitles.size()) {
            return false;
        }
        return sourceTitles.containsAll(targetTitles);
    }

    /**
     * 初始化SAX读取器
     * 将所有的数据按行读取
     *
     * @throws IOException
     */
    Excel07SaxReader initSaxReader(ExConsumer<R> dataConsumer) {

        HashMap<Integer, String> colIndex2Title = new HashMap<>();
        /*返回一个Sax读取器*/
        return new Excel07SaxReader(new RowHandler() {

            @Override
            public void doAfterAllAnalysed() {
                /*sheet读取结束时*/

            }

            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {
                /*读取标题*/
                if (sheetIndex == 0 && rowIndex == titleRowIndex) {
                    for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                        colIndex2Title.put(colIndex, MataConverter.parse(rowList.get(colIndex)));
                    }
                    return;
                }
                /*只读取第一个sheet的数据 只读取标题行之后的数据 */
                if (sheetIndex != 0 || rowIndex <= titleRowIndex) {
                    return;
                }
                /*Excel解析原生的数据*/
                HashMap<String, Object> rowData = new LinkedHashMap<>();
                for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                    rowData.put(colIndex2Title.get(colIndex), MataConverter.parse(rowList.get(colIndex)));
                }
                /*转换成业务对象*/
                try {
                    /*转换*/
                    rowConvert(rowData);
                    rowData.put(RESULT_TITLE, CONVERT_SUCCESS_MSG);
                    /*消费*/
                    dataConsumer.accept(currentObject);
                    rowData.put(RESULT_TITLE, IMPORT_SUCCESS_MSG);
                } catch (Exception e) {
                    /*写入导出Excel结果*/
                    rowData.put(RESULT_TITLE, GlobalExceptionManager.getExceptionMsg(e));
                } finally {
                    excelExport.writeByMap(Collections.singletonList(rowData));
                }
            }
        });
    }


    /**
     * 转换器策略
     */
    public enum ConvertStrategy {
        /*如果convert阶段失败, 跳过所有的行*/
        SKIP_ALL_IF_FAIL,
        /*如果convert阶段失败, 跳过失败的行*/
        SKIP_ROW_IF_FAIL;

    }

}
