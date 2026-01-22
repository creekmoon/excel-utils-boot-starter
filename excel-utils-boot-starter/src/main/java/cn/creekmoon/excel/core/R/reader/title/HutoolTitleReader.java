package cn.creekmoon.excel.core.R.reader.title;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.StringConverter;
import cn.creekmoon.excel.core.R.report.ImportOptions;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.exception.*;
import cn.hutool.core.text.StrFormatter;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelFileUtil;
import cn.hutool.poi.excel.sax.ExcelSaxReader;
import cn.hutool.poi.excel.sax.ExcelSaxUtil;
import cn.hutool.poi.excel.sax.StopReadException;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

import static cn.creekmoon.excel.util.ExcelConstants.*;

@Slf4j
public class HutoolTitleReader<R> extends TitleReader<R> {

    /*标志位, 模板一致性检查已经失败 */
    public boolean TEMPLATE_CONSISTENCY_CHECK_FAILED = false;

    public HutoolTitleReader(ExcelImport parent, String sheetRid, String sheetName, Supplier newObjectSupplier) {
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
    public <T> HutoolTitleReader<T> reset(Supplier<T> newObjectSupplier) {
        // 创建新的 reader 实例
        HutoolTitleReader<T> newReader = new HutoolTitleReader<>(
                this.getParent(),
                this.sheetRid,
                this.sheetName,
                newObjectSupplier
        );

        // 注意：不更新 ExcelImport 的 Map
        // 这样可以保留第一个（主）Reader 用于生成导入结果文件
        // Reset 创建的 Reader 只用于临时读取，不参与结果文件生成

        return newReader;
    }


    @Override
    public <T> HutoolTitleReader<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        super.title2converts.put(title, convert);
        super.title2consumers.put(title, setter);
        return this;
    }


    @Override
    public HutoolTitleReader<R> addConvert(String title, BiConsumer<R, String> reader) {
        addConvert(title, x -> x, reader);
        return this;
    }


    @Override
    public <T> HutoolTitleReader<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter) {
        super.skipEmptyTitles.add(title);
        addConvert(title, x -> x, setter);
        return this;
    }

    @Override
    public <T> HutoolTitleReader<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        super.skipEmptyTitles.add(title);
        addConvert(title, convert, setter);
        return this;
    }


    @Override
    public HutoolTitleReader<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter) {
        super.mustExistTitles.add(title);
        addConvert(title, x -> x, setter);
        return this;
    }

    @Override
    public <T> HutoolTitleReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
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
    public <T> HutoolTitleReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor) {
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
        // 初始化配置和报告
        super.options = options != null ? options : ImportOptions.defaults();
        initializeReport();
        
        // 执行读取（会填充 report）
        read((ExConsumer<R>) data -> {
            // 如果需要收集成功数据
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
        
        // 初始化错误列表（如果需要收集）
        if (super.options != null && super.options.isCollectRowErrors()) {
            super.report.setRowErrors(new ArrayList<>());
        }
    }

    @SneakyThrows
    @Override
    public TitleReader<R> read(ExBiConsumer<Integer, R> dataConsumer) {
        // 如果 report 尚未初始化，使用默认配置初始化
        if (super.report == null) {
            super.options = ImportOptions.defaults();
            initializeReport();
        }
        
        /*尝试拿锁*/
        ExcelUtilsConfig.importParallelSemaphore.acquire();
        try {
            //新版读取 使用SAX读取模式
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolTitleReader.read] 开始读取sheet: rId={}, sheetName={}",
                        sheetRid, sheetName);
            }

            ExcelSaxReader saxReader = initSaxReader(dataConsumer);
            /*第一个参数 文件流  第二个参数 sheetRid 直接定位到指定sheet*/
            saxReader.read(this.getParent().sourceFile.getInputStream(), getParent().rid2SheetNameBiMap.get(sheetRid));

            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolTitleReader.read] Sheet读取完成: rId={}, sheetName={}",
                        sheetRid, sheetName);
            }
        } catch (StopReadException e) {
            // StopReadException 是正常的提前终止，不算异常
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolTitleReader.read] 读取提前终止（fail-fast或阈值停止）: rId={}, sheetName={}",
                        sheetRid, sheetName);
            }
        } catch (Exception e) {
            log.error("SaxReader读取Excel文件异常, sheetRid={}, sheetName={}", sheetRid, sheetName, e);
            // 记录到 report
            if (super.report != null && !GlobalExceptionMsgManager.isCustomException((Exception) e)) {
                super.report.setGlobalErrorCode("FATAL");
                super.report.setGlobalErrorMessage(GlobalExceptionMsgManager.getExceptionMsg((Exception) e));
            }
        } finally {
            /*释放信号量*/
            ExcelUtilsConfig.importParallelSemaphore.release();
        }
        return this;
    }

    /**
     * 增加读取范围限制
     *
     * @param titleRowIndex    标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param lastDataRowIndex 最后一条数据所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @return
     */
    @Override
    public HutoolTitleReader<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex) {
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
    public HutoolTitleReader<R> range(int startRowIndex, int lastRowIndex) {
        return range(startRowIndex, startRowIndex + 1, lastRowIndex);
    }

    /**
     * 增加读取范围限制
     *
     * @param startRowIndex 起始行下标(从0开始)
     * @return
     */
    @Override
    public HutoolTitleReader<R> range(int startRowIndex) {
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
            throw new CheckedExcelException(TITLE_CHECK_ERROR);
        }

        /*过滤空白行*/
        if (super.ENABLE_BLANK_ROW_FILTER
                && row.values().stream().allMatch(x -> x == null || "".equals(x))
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
            String value = Optional.ofNullable(entry.getValue()).map(x -> (String) x).orElse("");
            /*检查必填项/检查可填项*/
            if (StrUtil.isBlank(value)) {
                if (super.mustExistTitles.contains(entry.getKey())) {
                    throw new CheckedExcelException(StrFormatter.format(FIELD_LACK_MSG, entry.getKey()));
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
                throw new CheckedExcelException(StrFormatter.format(ExcelConstants.CONVERT_FAIL_MSG + GlobalExceptionMsgManager.getExceptionMsg(e), entry.getKey()));
            }
        }
        return convertObject;
    }

    /**
     * 初始化SAX读取器
     *
     * @param
     * @return
     */

    ExcelSaxReader initSaxReader(ExBiConsumer<Integer, R> dataConsumer) throws IOException {

        /*返回一个Sax读取器*/
        return ExcelSaxUtil.createSaxReader(ExcelFileUtil.isXlsx(getParent().sourceFile.getInputStream()), new RowHandler() {

            @Override
            public void doAfterAllAnalysed() {
                /*sheet读取结束时*/
            }


            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {
                // 由于已通过rId精准定位sheet，无需在回调中过滤

                /*读取标题*/
                if (rowIndex == titleRowIndex) {
                    for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                        colIndex2Title.put(colIndex, StringConverter.parse(rowList.get(colIndex)));
                    }
                    
                    // 标题行读取完成后，立即进行模板一致性检查
                    if (ENABLE_TEMPLATE_CONSISTENCY_REVIEW) {
                        Set<String> targetTitles = title2converts.keySet();
                        Set<String> sourceTitles = new HashSet<>(colIndex2Title.values());
                        
                        if (!templateConsistencyCheck(targetTitles, sourceTitles)) {
                            // 模板不匹配
                            report.setGlobalErrorCode("TEMPLATE_MISMATCH");
                            report.setGlobalErrorMessage(TITLE_CHECK_ERROR);
                            
                            // 如果配置了 fail-fast，立即终止
                            if (options != null && options.isFailFastOnTemplateMismatch()) {
                                throw new StopReadException();
                            }
                            
                            // 否则标记为已失败，后续每行都会因此失败
                            TEMPLATE_CONSISTENCY_CHECK_FAILED = true;
                        }
                        
                        // 检查完成后关闭，避免重复检查
                        ENABLE_TEMPLATE_CONSISTENCY_REVIEW = false;
                    }
                    
                    return;
                }
                /*只读取指定范围的数据 */
                if (rowIndex == (int) titleRowIndex
                        || rowIndex < firstRowIndex
                        || rowIndex > latestRowIndex) {
                    return;
                }
                /*没有添加 convert直接跳过 */
                if (title2converts.isEmpty()
                        && title2consumers.isEmpty()
                ) {
                    return;
                }

                /*Excel解析原生的数据. 目前只用于内部数据转换*/
                HashMap<String, String> hashDataMap = new LinkedHashMap<>();
                for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                    hashDataMap.put(colIndex2Title.get(colIndex), StringConverter.parse(rowList.get(colIndex)));
                }
                /*转换成业务对象*/
                R currentObject = null;
                try {
                    /*转换*/
                    currentObject = rowConvert(hashDataMap);
                    if (currentObject == null) {
                        // 空白行，不计入统计
                        return;
                    }
                    
                    // 更新 processed 范围
                    if (report.getProcessedFirstRowIndex() == null) {
                        report.setProcessedFirstRowIndex((int) rowIndex);
                    }
                    report.setProcessedLastRowIndex((int) rowIndex);
                    report.setTotalRows(report.getTotalRows() + 1);
                    
                    /*转换后置处理器*/
                    for (ExConsumer convertPostProcessor : convertPostProcessors) {
                        convertPostProcessor.accept(currentObject);
                    }
                    
                    // 成功：更新统计
                    report.setSuccessRows(report.getSuccessRows() + 1);
                    dataConsumer.accept((int) rowIndex, currentObject);
                    
                } catch (Exception e) {
                    //先记录异常信息
                    EXISTS_READ_FAIL.set(true);
                    String exceptionMsg = GlobalExceptionMsgManager.getExceptionMsg(e);
                    
                    // 更新 processed 范围（失败行也算处理过）
                    if (report.getProcessedFirstRowIndex() == null) {
                        report.setProcessedFirstRowIndex((int) rowIndex);
                    }
                    report.setProcessedLastRowIndex((int) rowIndex);
                    report.setTotalRows(report.getTotalRows() + 1);
                    
                    // 失败：更新统计和错误明细
                    report.setErrorRows(report.getErrorRows() + 1);
                    
                    // 收集错误明细
                    if (options != null && options.isCollectRowErrors()) {
                        String errorCode = determineErrorCode(e);
                        report.getRowErrors().add(new ImportReport.RowError((int) rowIndex, errorCode, exceptionMsg));
                    }
                    
                    // 检查错误阈值
                    if (options != null && options.getMaxErrorRows() != null 
                            && report.getErrorRows() >= options.getMaxErrorRows()) {
                        report.setGlobalErrorCode("ERROR_LIMIT_REACHED");
                        report.setGlobalErrorMessage("错误行数达到阈值 " + options.getMaxErrorRows() + "，已停止读取");
                        throw new StopReadException();
                    }
                    
                    //如果是非自定义异常,中断读取
                    if (!GlobalExceptionMsgManager.isCustomException(e)) {
                        report.setGlobalErrorCode("FATAL");
                        report.setGlobalErrorMessage(exceptionMsg);
                        throw new RuntimeException(e);
                    }
                }
            }
        });
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
    public HutoolTitleReader<R> disableTemplateConsistencyCheck() {
        super.ENABLE_TEMPLATE_CONSISTENCY_REVIEW = false;
        return this;
    }

    /**
     * 禁用空白行过滤
     *
     * @return
     */
    public HutoolTitleReader<R> disableBlankRowFilter() {
        super.ENABLE_BLANK_ROW_FILTER = false;
        return this;
    }

    /**
     * 根据异常类型确定错误码
     */
    private String determineErrorCode(Exception e) {
        if (e instanceof CheckedExcelException) {
            String msg = e.getMessage();
            if (msg != null && msg.contains(TITLE_CHECK_ERROR)) {
                return "TEMPLATE_MISMATCH";
            }
            return "VALIDATION_ERROR";
        }
        return "FATAL";
    }
}
