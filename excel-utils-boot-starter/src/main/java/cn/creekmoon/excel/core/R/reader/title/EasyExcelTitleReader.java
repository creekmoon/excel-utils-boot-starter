package cn.creekmoon.excel.core.R.reader.title;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.StringConverter;
import cn.creekmoon.excel.core.R.report.ImportGlobalErrorEnum;
import cn.creekmoon.excel.core.R.report.ImportOptions;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.creekmoon.excel.util.exception.ExBiConsumer;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;
import cn.creekmoon.excel.util.exception.GlobalExceptionMsgManager;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.ReadDefaultReturnEnum;
import com.alibaba.excel.event.AnalysisEventListener;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

@Slf4j
public class EasyExcelTitleReader<R> extends TitleReader<R> {

    /*标志位, 模板一致性检查已经失败 */
    public boolean TEMPLATE_CONSISTENCY_CHECK_FAILED = false;

    public EasyExcelTitleReader(ExcelImport parent, String sheetRid, String sheetName, Supplier newObjectSupplier) {
        super(parent);
        super.sheetRid = sheetRid;
        super.sheetName = sheetName;
        super.newObjectSupplier = newObjectSupplier;
    }

    /**
     * 重置读取器以支持在同一个sheet中读取不同类型的表格
     * 新的读取器会清空所有转换规则和范围设置
     * 需要重新调用 addConvert() 和 range() 配置
     * <p>
     * 重要限制：
     * - Reset 创建的 Reader 不会参与 ExcelImport.generateResultFile() 的结果生成
     * - 如果需要完整的导入验证结果报告，建议为每个数据类型使用独立的 switchSheet()
     * - Reset 适用于在同一 Sheet 中读取多个数据区域，但不需要生成统一验证报告的场景
     *
     * @param newObjectSupplier 新表格的对象创建器
     * @param <T>               新的数据类型
     * @return 新的 TitleReader 实例
     */
    @Override
    public <T> EasyExcelTitleReader<T> reset(Supplier<T> newObjectSupplier) {
        EasyExcelTitleReader<T> newReader = new EasyExcelTitleReader<>(
                this.getParent(),
                this.sheetRid,
                this.sheetName,
                newObjectSupplier
        );
        return newReader;
    }

    @Override
    public <T> EasyExcelTitleReader<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        super.title2converts.put(title, convert);
        super.title2consumers.put(title, setter);
        return this;
    }

    @Override
    public EasyExcelTitleReader<R> addConvert(String title, BiConsumer<R, String> reader) {
        addConvert(title, x -> x, reader);
        return this;
    }

    @Override
    public <T> EasyExcelTitleReader<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter) {
        super.skipEmptyTitles.add(title);
        addConvert(title, x -> x, setter);
        return this;
    }

    @Override
    public <T> EasyExcelTitleReader<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        super.skipEmptyTitles.add(title);
        addConvert(title, convert, setter);
        return this;
    }

    @Override
    public EasyExcelTitleReader<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter) {
        super.mustExistTitles.add(title);
        addConvert(title, x -> x, setter);
        return this;
    }

    @Override
    public <T> EasyExcelTitleReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        super.mustExistTitles.add(title);
        addConvert(title, convert, setter);
        return this;
    }

    /**
     * 添加校验阶段后置处理器 当所有的convert执行完成后会执行这个操作做最后的校验处理
     *
     * @param postProcessor 后置处理器
     * @param <T>
     * @return
     */
    @Override
    public <T> EasyExcelTitleReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor) {
        if (postProcessor != null) {
            super.convertPostProcessors.add(postProcessor);
        }
        return this;
    }

    @Override
    public TitleReader<R> read(ExConsumer<R> dataConsumer) {
        return read((Integer rowIndex, R data) -> dataConsumer.accept(data));
    }

    @Override
    public ImportReport<R> readWithReport(ImportOptions options) {
        super.options = options != null ? options : ImportOptions.defaults();
        initializeReport();

        read((ExConsumer<R>) data -> {
            if (super.options.isCollectSuccessData()) {
                if (super.report.getData() == null) {
                    super.report.setData(new ArrayList<>());
                }
                Integer maxSuccessData = super.options.getMaxSuccessData();
                if (maxSuccessData == null || super.report.getData().size() < maxSuccessData) {
                    super.report.getData().add(data);
                }
            }
        });

        return super.report;
    }

    /**
     * 初始化导入报告
     */
    private void initializeReport() {
        super.report = new ImportReport<>();
        super.report.setTaskId(getParent().taskId);
        super.report.setSheetRid(sheetRid);
        super.report.setSheetName(sheetName);
        super.report.setTitleRowIndex(titleRowIndex);
        super.report.setFirstRowIndex(firstRowIndex);
        super.report.setLatestRowIndex(latestRowIndex);

        if (super.options != null && super.options.isCollectRowErrors()) {
            super.report.setRowErrors(new ArrayList<>());
        }
    }

    @SneakyThrows
    @Override
    public TitleReader<R> read(ExBiConsumer<Integer, R> dataConsumer) {
        if (super.report == null) {
            super.options = ImportOptions.defaults();
            initializeReport();
        }

        ExcelUtilsConfig.importParallelSemaphore.acquire();
        try {
            log.debug("[EasyExcelTitleReader] 开始读取sheet: rId={}, sheetName={}",
                    sheetRid, sheetName);

            EasyExcel.read(this.getParent().sourceFile.getInputStream(), new AnalysisEventListener<Map<Integer, String>>() {
                @Override
                public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
                    int rowIndex = context.readRowHolder().getRowIndex();
                    log.debug("[EasyExcelTitleReader] invokeHeadMap rowIndex={}, headMap={}", rowIndex, headMap);
                    if (rowIndex == titleRowIndex) {
                        for (Map.Entry<Integer, String> entry : headMap.entrySet()) {
                            colIndex2Title.put(entry.getKey(), StringConverter.parse(entry.getValue()));
                        }

                        if (ENABLE_TEMPLATE_CONSISTENCY_REVIEW) {
                            Set<String> targetTitles = title2converts.keySet();
                            Set<String> sourceTitles = new HashSet<>(colIndex2Title.values());

                            if (!templateConsistencyCheck(targetTitles, sourceTitles)) {
                                addGlobalError(ImportGlobalErrorEnum.TEMPLATE_MISMATCH);

                                if (options != null && options.isFailFastOnTemplateMismatch()) {
                                    throw new EasyExcelStopReadException();
                                }

                                TEMPLATE_CONSISTENCY_CHECK_FAILED = true;
                            }

                            ENABLE_TEMPLATE_CONSISTENCY_REVIEW = false;
                        }
                    }
                }

                @Override
                public void invoke(Map<Integer, String> data, AnalysisContext context) {
                    int rowIndex = context.readRowHolder().getRowIndex();
                    log.debug("[EasyExcelTitleReader] invoke rowIndex={}, dataSize={}, titleRowIndex={}, firstRowIndex={}, latestRowIndex={}",
                            rowIndex, data.size(), titleRowIndex, firstRowIndex, latestRowIndex);

                    /*只读取指定范围的数据 */
                    if (rowIndex < firstRowIndex || rowIndex > latestRowIndex) {
                        log.debug("[EasyExcelTitleReader] row {} out of range [{}, {}]", rowIndex, firstRowIndex, latestRowIndex);
                        return;
                    }

                    /*没有添加 convert直接跳过 */
                    if (title2converts.isEmpty() && title2consumers.isEmpty()) {
                        log.debug("[EasyExcelTitleReader] no converts registered");
                        return;
                    }
                    log.debug("[EasyExcelTitleReader] processing data row {}, colIndex2Title size={}", rowIndex, colIndex2Title.size());

                    /*Excel解析原生的数据. 目前只用于内部数据转换*/
                    HashMap<String, String> hashDataMap = new LinkedHashMap<>();
                    for (Map.Entry<Integer, String> entry : colIndex2Title.entrySet()) {
                        Integer colIndex = entry.getKey();
                        String title = entry.getValue();
                        Object cellValue = data.get(colIndex);
                        hashDataMap.put(title, StringConverter.parse(cellValue));
                    }

                    /*转换成业务对象*/
                    R currentObject = null;
                    boolean totalRowsCounted = false;
                    try {
                        currentObject = rowConvert(hashDataMap);
                        if (currentObject == null) {
                            return;
                        }

                        if (report.getProcessedFirstRowIndex() == null) {
                            report.setProcessedFirstRowIndex(rowIndex);
                        }
                        report.setProcessedLastRowIndex(rowIndex);
                        report.setTotalRows(report.getTotalRows() + 1);
                        totalRowsCounted = true;

                        /*转换后置处理器*/
                        for (ExConsumer<R> convertPostProcessor : convertPostProcessors) {
                            convertPostProcessor.accept(currentObject);
                        }

                        dataConsumer.accept(rowIndex, currentObject);

                        report.setSuccessRows(report.getSuccessRows() + 1);

                    } catch (Exception e) {
                        EXISTS_READ_FAIL.set(true);
                        String exceptionMsg = GlobalExceptionMsgManager.getExceptionMsg(e);

                        if (!totalRowsCounted) {
                            if (report.getProcessedFirstRowIndex() == null) {
                                report.setProcessedFirstRowIndex(rowIndex);
                            }
                            report.setProcessedLastRowIndex(rowIndex);
                            report.setTotalRows(report.getTotalRows() + 1);
                        }

                        report.setErrorRows(report.getErrorRows() + 1);

                        if (options != null && options.isCollectRowErrors() && !isTemplateMismatchException(e)) {
                            report.getRowErrors().add(new ImportReport.RowError(rowIndex, exceptionMsg));
                        }

                        if (options != null && options.getMaxErrorRows() != null
                                && report.getErrorRows() >= options.getMaxErrorRows()) {
                            addGlobalError(ImportGlobalErrorEnum.ERROR_LIMIT_REACHED);
                            throw new EasyExcelStopReadException();
                        }

                        if (!GlobalExceptionMsgManager.isCustomException(e)) {
                            addGlobalError(ImportGlobalErrorEnum.IMPORT_INTERRUPTED);
                            throw new RuntimeException(e);
                        }
                    }
                }

                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    /*sheet读取结束时*/
                }
            }).readDefaultReturn(ReadDefaultReturnEnum.ACTUAL_DATA)
                    /* Map 模式返回真实 Java 类型（日期→Date），避免默认 STRING 模式把日期
                     * 转成分钟精度字符串导致秒级丢失；Date 由 StringConverter 统一格式化 */
                    .sheet(sheetName)
                    /* 表头行数 = titleRowIndex + 1：invokeHeadMap 只对 EasyExcel 认定的表头行触发，
                     * 非 0 表头（多级表头/同 sheet 多表格）必须扩大表头行范围 */
                    .headRowNumber(titleRowIndex + 1)
                    .doRead();

            log.debug("[EasyExcelTitleReader] Sheet读取完成: rId={}, sheetName={}",
                    sheetRid, sheetName);
        } catch (EasyExcelStopReadException e) {
            log.debug("[EasyExcelTitleReader] 读取提前终止（fail-fast或阈值停止）: rId={}, sheetName={}",
                    sheetRid, sheetName);
        } catch (Exception e) {
            log.error("EasyExcel读取Excel文件异常, sheetRid={}, sheetName={}, exClass={}, exMsg={}",
                    sheetRid, sheetName, e.getClass().getName(), e.getMessage(), e);
            if (super.report != null && !GlobalExceptionMsgManager.isCustomException(e)) {
                addGlobalError(ImportGlobalErrorEnum.IMPORT_INTERRUPTED);
            }
        } finally {
            ExcelUtilsConfig.importParallelSemaphore.release();
        }
        return this;
    }

    /**
     * 增加读取范围限制
     *
     * @param titleRowIndex    标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param firstDataRowIndex 首条数据所在的行数(下标按照从0开始)
     * @param lastDataRowIndex 最后一条数据所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @return
     */
    @Override
    public EasyExcelTitleReader<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex) {
        super.titleRowIndex = titleRowIndex;
        super.firstRowIndex = firstDataRowIndex;
        super.latestRowIndex = lastDataRowIndex;
        return this;
    }

    /**
     * 增加读取范围限制
     *
     * @param startRowIndex 标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param lastRowIndex  最后一条数据所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @return
     */
    @Override
    public EasyExcelTitleReader<R> range(int startRowIndex, int lastRowIndex) {
        return range(startRowIndex, startRowIndex + 1, lastRowIndex);
    }

    /**
     * 增加读取范围限制
     *
     * @param startRowIndex 起始行下标(从0开始)
     * @return
     */
    @Override
    public EasyExcelTitleReader<R> range(int startRowIndex) {
        return range(startRowIndex, startRowIndex + 1, Integer.MAX_VALUE);
    }

    /**
     * 行转换
     *
     * @param row 实际上是Map<String, String>对象
     * @throws Exception
     */
    private R rowConvert(Map<String, String> row) throws Exception {
        /*模板一致性检查已前置到标题行，这里检查标记即可*/
        if (TEMPLATE_CONSISTENCY_CHECK_FAILED) {
            throw new CheckedExcelException(ImportGlobalErrorEnum.TEMPLATE_MISMATCH.getMessage());
        }

        /*过滤空白行*/
        if (super.ENABLE_BLANK_ROW_FILTER
                && row.values().stream().allMatch(x -> x == null || x.isEmpty())
        ) {
            return null;
        }

        /*初始化空对象*/
        R convertObject = (R) super.newObjectSupplier.get();
        /*最大转换次数*/
        int maxConvertCount = super.title2consumers.keySet().size();
        /*执行convert*/
        for (Map.Entry<String, String> entry : row.entrySet()) {
            /*如果包含不支持的标题,  或者已经超过最大次数则不再进行读取*/
            if (!super.title2consumers.containsKey(entry.getKey()) || maxConvertCount-- <= 0) {
                continue;
            }
            String value = Optional.ofNullable(entry.getValue()).orElse("");
            /*检查必填项/检查可填项*/
            if (value == null || value.trim().isEmpty()) {
                if (super.mustExistTitles.contains(entry.getKey())) {
                    throw new CheckedExcelException(String.format(ExcelConstants.FIELD_LACK_MSG, entry.getKey()));
                }
                if (super.skipEmptyTitles.contains(entry.getKey())) {
                    continue;
                }
            }
            /*转换数据*/
            try {
                Object convertValue = super.title2converts.get(entry.getKey()).apply(value);
                super.title2consumers.get(entry.getKey()).accept(convertObject, convertValue);
            } catch (Exception e) {
                log.warn("EXCEL导入数据转换失败！", e);
                throw new CheckedExcelException(String.format(ExcelConstants.CONVERT_FAIL_MSG + GlobalExceptionMsgManager.getExceptionMsg(e), entry.getKey()));
            }
        }
        return convertObject;
    }

    /**
     * 模版一致性检查
     *
     * @param targetTitles 我们声明的要拿取的标题
     * @param sourceTitles 传过来的excel文件标题
     * @return
     */
    private Boolean templateConsistencyCheck(Set<String> targetTitles, Set<String> sourceTitles) {
        if (targetTitles.size() > sourceTitles.size()) {
            return false;
        }
        return sourceTitles.containsAll(targetTitles);
    }

    /**
     * 禁用模版一致性检查
     *
     * @return
     */
    public EasyExcelTitleReader<R> disableTemplateConsistencyCheck() {
        super.ENABLE_TEMPLATE_CONSISTENCY_REVIEW = false;
        return this;
    }

    /**
     * 禁用空白行过滤
     *
     * @return
     */
    public EasyExcelTitleReader<R> disableBlankRowFilter() {
        super.ENABLE_BLANK_ROW_FILTER = false;
        return this;
    }

    private boolean isTemplateMismatchException(Exception e) {
        return e instanceof CheckedExcelException
                && e.getMessage() != null
                && e.getMessage().equals(ImportGlobalErrorEnum.TEMPLATE_MISMATCH.getMessage());
    }

    private void addGlobalError(ImportGlobalErrorEnum globalError) {
        if (report.getGlobalErrors() == null) {
            report.setGlobalErrors(new LinkedHashSet<>());
        }
        report.getGlobalErrors().add(globalError);
    }

    /**
     * 用于在 EasyExcel 读取过程中提前终止读取的内部异常
     */
    private static class EasyExcelStopReadException extends RuntimeException {
    }
}
