package cn.creekmoon.excel.core.R.reader.title;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.StringConverter;
import cn.creekmoon.excel.core.R.reader.SheetIndexNormalizingRowHandler;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;
import cn.creekmoon.excel.util.exception.GlobalExceptionMsgManager;
import cn.hutool.core.text.StrFormatter;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelFileUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.ExcelSaxReader;
import cn.hutool.poi.excel.sax.ExcelSaxUtil;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

import static cn.creekmoon.excel.util.ExcelConstants.*;

@Slf4j
public class HutoolTitleReader<R> extends TitleReader<R> {


    public HutoolTitleReader(ExcelImport parent, Integer sheetIndex, Supplier newObjectSupplier) {
        super(parent);
        super.readerResult = new TitleReaderResult<R>();
        super.sheetIndex = sheetIndex;
        super.newObjectSupplier = newObjectSupplier;
    }

    /**
     * 重置读取器以支持在同一个sheet中读取不同类型的表格
     * 新的读取器会清空所有转换规则和范围设置
     * 需要重新调用 addConvert() 和 range() 配置
     * 
     * 重要限制：
     * - Reset 创建的 Reader 不会参与 ExcelImport.generateResultFile() 的结果生成
     * - 如果需要完整的导入验证结果报告，建议为每个数据类型使用独立的 switchSheet()
     * - Reset 适用于在同一 Sheet 中读取多个数据区域，但不需要生成统一验证报告的场景
     *
     * @param newObjectSupplier 新表格的对象创建器
     * @param <T> 新的数据类型
     * @return 新的 TitleReader 实例
     */
    @Override
    public <T> HutoolTitleReader<T> reset(Supplier<T> newObjectSupplier) {
        // 创建新的 reader 实例
        HutoolTitleReader<T> newReader = new HutoolTitleReader<>(
            this.getParent(), 
            this.sheetIndex, 
            newObjectSupplier
        );
        
        // 注意：不更新 ExcelImport 的 Map
        // 这样可以保留第一个（主）Reader 用于生成导入结果文件
        // Reset 创建的 Reader 只用于临时读取，不参与结果文件生成
        
        return newReader;
    }

    /**
     * 获取SHEET页的总行数
     *
     * @return
     */
    @SneakyThrows
    @Override
    public Long getSheetRowCount() {
        if (getParent().debugger) {
            log.info("[DEBUGGER][HutoolTitleReader.getSheetRowCount] 开始统计行数: sheetIndex={}", super.sheetIndex);
        }
        
        AtomicLong result = new AtomicLong(0);
        
        // 获取sheetIndex对应的rid
        String rid = getParent().getSheetRid(super.sheetIndex);
        if (rid == null || rid.isEmpty()) {
            rid = "-1";
            if (getParent().debugger) {
                log.warn("[DEBUGGER][HutoolTitleReader.getSheetRowCount] 无法获取rid，降级使用-1");
            }
        }
        
        /*创建计数RowHandler*/
        RowHandler countHandler = new RowHandler() {
            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowCells) {
                if (sheetIndex != HutoolTitleReader.super.sheetIndex) {
                    return;
                }
                result.incrementAndGet();
            }
        };
        
        /*用适配器包装*/
        boolean isSingleSheetMode = !"-1".equals(rid);
        RowHandler wrappedHandler = new SheetIndexNormalizingRowHandler(
            countHandler,
            super.sheetIndex,
            isSingleSheetMode,
            getParent().debugger
        );
        
        /*创建SaxReader*/
        ExcelSaxReader<?> excel07SaxReader = ExcelSaxUtil.createSaxReader(
            ExcelFileUtil.isXlsx(getParent().sourceFile.getInputStream()),
            wrappedHandler
        );
        
        try {
            /*使用rid读取*/
            excel07SaxReader.read(getParent().sourceFile.getInputStream(), rid);
        } catch (Exception e) {
            log.error("getSheetRowCount方法读取文件异常", e);
        }
        
        if (getParent().debugger) {
            log.info("[DEBUGGER][HutoolTitleReader.getSheetRowCount] 统计完成: sheetIndex={}, rowCount={}", 
                super.sheetIndex, result.get());
        }
        
        return result.get();
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
    public TitleReaderResult<R> read(ExConsumer<R> dataConsumer) {
        TitleReaderResult<R> result = read().consume(dataConsumer);
        ((TitleReaderResult) getReadResult()).consumeSuccessTime = LocalDateTime.now();
        return result;
    }

    @SneakyThrows
    @Override
    public TitleReaderResult<R> read() {
        if (getParent().debugger) {
            log.info("[DEBUGGER][HutoolTitleReader.read] 开始读取: 目标sheetIndex={}, 文件名={}", 
                super.sheetIndex, getParent().sourceFile.getOriginalFilename());
        }
        
        /*尝试拿锁*/
        ExcelUtilsConfig.importParallelSemaphore.acquire();
        getReadResult().readStartTime = LocalDateTime.now();
        try {
            //新版读取 使用SAX读取模式

            // 获取sheetIndex对应的rid
            String rid = getParent().getSheetRid(super.sheetIndex);
            
            // 确定是否使用单sheet优化模式
            boolean isSingleSheetMode;
            if (rid == null || rid.isEmpty()) {
                // 映射失败，降级使用-1（读取所有sheet）
                rid = "-1";
                isSingleSheetMode = false;
                if (getParent().debugger) {
                    log.warn("[DEBUGGER][HutoolTitleReader.read] 无法获取rid，降级使用-1: sheetIndex={}", super.sheetIndex);
                }
            } else {
                // 使用单sheet优化
                isSingleSheetMode = true;
                if (getParent().debugger) {
                    log.info("[DEBUGGER][HutoolTitleReader.read] 使用单sheet优化: sheetIndex={} → rid={}", 
                        super.sheetIndex, rid);
                }
            }
            
            /*创建原始RowHandler*/
            RowHandler originalHandler = createRowHandler();
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolTitleReader.read] 创建原始RowHandler成功");
            }
            
            /*用适配器包装*/
            RowHandler wrappedHandler = new SheetIndexNormalizingRowHandler(
                originalHandler, 
                super.sheetIndex, 
                isSingleSheetMode,
                getParent().debugger
            );
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolTitleReader.read] 创建适配器包装成功: actualSheetIndex={}, singleSheetMode={}", 
                    super.sheetIndex, isSingleSheetMode);
            }
            
            /*创建SaxReader*/
            ExcelSaxReader saxReader = createSaxReader(wrappedHandler);
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolTitleReader.read] 准备调用saxReader.read(): rid={}, saxReaderClass={}", 
                    rid, saxReader.getClass().getName());
                log.info("[DEBUGGER][HutoolTitleReader.read] 文件流信息: available={} bytes", 
                    getParent().sourceFile.getInputStream().available());
            }
            
            /*使用rid读取（可能是单sheet或所有sheet）*/
            saxReader.read(this.getParent().sourceFile.getInputStream(), rid);
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolTitleReader.read] saxReader.read()执行完成");
            }
        } catch (Exception e) {
            log.error("SaxReader读取Excel文件异常", e);
            if (getParent().debugger) {
                log.error("[DEBUGGER][HutoolTitleReader.read] 读取异常: {}", e.getMessage(), e);
            }
        } finally {
            getReadResult().readSuccessTime = LocalDateTime.now();
            /*释放信号量*/
            ExcelUtilsConfig.importParallelSemaphore.release();
            
            if (getParent().debugger) {
                log.info("[DEBUGGER][HutoolTitleReader.read] 读取完成: 耗时={}ms", 
                    java.time.Duration.between(getReadResult().readStartTime, getReadResult().readSuccessTime).toMillis());
            }
        }
        return getReadResult();
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
        /*进行模板一致性检查*/
        if (super.TEMPLATE_CONSISTENCY_CHECK_ENABLE) {
            if (super.TEMPLATE_CONSISTENCY_CHECK_FAILED || !templateConsistencyCheck(super.title2converts.keySet(), row.keySet())) {
                super.TEMPLATE_CONSISTENCY_CHECK_FAILED = true;
                throw new CheckedExcelException(TITLE_CHECK_ERROR);
            }
        }
        super.TEMPLATE_CONSISTENCY_CHECK_ENABLE = false;

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
     * 创建RowHandler（业务处理逻辑）
     *
     * @return RowHandler实例
     */
    RowHandler createRowHandler() {
        int targetSheetIndex = super.sheetIndex;
        TitleReaderResult titleReaderResult = (TitleReaderResult) getReadResult();

        return new RowHandler() {

            @Override
            public void doAfterAllAnalysed() {
                /*sheet读取结束时*/
                if (getParent().debugger) {
                    log.info("[DEBUGGER][HutoolTitleReader.doAfterAllAnalysed] Sheet读取结束: targetSheetIndex={}, colIndex2Title.size={}, 总数据行数={}", 
                        targetSheetIndex, colIndex2Title.size(), titleReaderResult.rowIndex2dataBiMap.size());
                }
            }


            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {
                if (getParent().debugger) {
                    log.info("[DEBUGGER][HutoolTitleReader.handle] 接收回调: sheetIndex={}, rowIndex={}, rowListSize={}, targetSheetIndex={}", 
                        sheetIndex, rowIndex, rowList == null ? 0 : rowList.size(), targetSheetIndex);
                    // 🔍 打印前几行的实际内容以确认sheet是否正确
                    if (rowIndex <= 5) {
                        log.info("[DEBUGGER][HutoolTitleReader.handle] 🔍实际行内容(前6行): rowIndex={}, rowList={}", 
                            rowIndex, rowList);
                    }
                }
                
                if (targetSheetIndex != sheetIndex) {
                    if (getParent().debugger) {
                        log.info("[DEBUGGER][HutoolTitleReader.handle] Sheet过滤: targetSheetIndex({}) != sheetIndex({}), 跳过此行", 
                            targetSheetIndex, sheetIndex);
                    }
                    return;
                }

                /*读取标题*/
                if (getParent().debugger) {
                    log.info("[DEBUGGER][HutoolTitleReader.handle] 标题行检查: rowIndex={}, titleRowIndex={}, rowIndex==titleRowIndex={}", 
                        rowIndex, titleRowIndex, rowIndex == titleRowIndex);
                }
                if (rowIndex == titleRowIndex) {
                    if (getParent().debugger) {
                        log.info("[DEBUGGER][HutoolTitleReader.handle] 进入标题行读取逻辑: rowListSize={}, rowList={}", 
                            rowList.size(), rowList);
                    }
                    for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                        colIndex2Title.put(colIndex, StringConverter.parse(rowList.get(colIndex)));
                    }
                    if (getParent().debugger) {
                        log.info("[DEBUGGER][HutoolTitleReader.handle] 读取标题行完成: rowIndex={}, 标题={}", 
                            rowIndex, colIndex2Title);
                    }
                    return;
                }
                /*只读取指定范围的数据 */
                if (rowIndex == (int) titleRowIndex
                        || rowIndex < firstRowIndex
                        || rowIndex > latestRowIndex) {
                    if (getParent().debugger) {
                        log.info("[DEBUGGER][HutoolTitleReader.handle] 范围过滤: rowIndex={}, titleRowIndex={}, firstRowIndex={}, latestRowIndex={}, 跳过此行", 
                            rowIndex, titleRowIndex, firstRowIndex, latestRowIndex);
                    }
                    return;
                }
                /*没有添加 convert直接跳过 */
                if (title2converts.isEmpty()
                        && title2consumers.isEmpty()
                ) {
                    if (getParent().debugger) {
                        log.info("[DEBUGGER][HutoolTitleReader.handle] Convert配置为空: title2converts.size={}, title2consumers.size={}, 跳过此行", 
                            title2converts.size(), title2consumers.size());
                    }
                    return;
                }

                /*Excel解析原生的数据. 目前只用于内部数据转换*/
                if (getParent().debugger) {
                    log.info("[DEBUGGER][HutoolTitleReader.handle] 准备构建hashDataMap: rowIndex={}, colIndex2Title={}, rowListSize={}", 
                        rowIndex, colIndex2Title, rowList.size());
                }
                HashMap<String, String> hashDataMap = new LinkedHashMap<>();
                for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                    String titleKey = colIndex2Title.get(colIndex);
                    String cellValue = StringConverter.parse(rowList.get(colIndex));
                    if (getParent().debugger && colIndex < 3) {
                        log.info("[DEBUGGER][HutoolTitleReader.handle] 映射单元格: colIndex={}, title={}, value={}", 
                            colIndex, titleKey, cellValue);
                    }
                    hashDataMap.put(titleKey, cellValue);
                }
                
                if (getParent().debugger) {
                    log.info("[DEBUGGER][HutoolTitleReader.handle] 构建hashDataMap成功: rowIndex={}, dataMap.keys={}", 
                        rowIndex, hashDataMap.keySet());
                }
                
                /*转换成业务对象*/
                R currentObject = null;
                try {
                    /*转换*/
                    currentObject = rowConvert(hashDataMap);
                    if (currentObject == null) {
                        if (getParent().debugger) {
                            log.info("[DEBUGGER][HutoolTitleReader.handle] rowConvert返回null (可能是空白行), 跳过");
                        }
                        return;
                    }
                    /*转换后置处理器*/
                    for (ExConsumer convertPostProcessor : convertPostProcessors) {
                        convertPostProcessor.accept(currentObject);
                    }
                    titleReaderResult.rowIndex2msg.put((int) rowIndex, CONVERT_SUCCESS_MSG);
                    /*消费 - 根据enableDataMemoryCache决定是否缓存数据对象*/
                    if (enableDataMemoryCache) {
                        titleReaderResult.rowIndex2dataBiMap.put((int) rowIndex, currentObject);
                    }
                    
                    if (getParent().debugger) {
                        log.info("[DEBUGGER][HutoolTitleReader.handle] 数据处理成功: rowIndex={}, object={}", 
                            rowIndex, currentObject);
                    }
                } catch (Exception e) {
                    //假如存在任一数据convert阶段就失败的单, 将打一个标记
                    titleReaderResult.EXISTS_READ_FAIL.set(true);
                    titleReaderResult.errorCount.incrementAndGet();
                    /*写入导出Excel结果*/
                    String exceptionMsg = GlobalExceptionMsgManager.getExceptionMsg(e);
                    getReadResult().errorReport.append(StrFormatter.format("第[{}]行发生错误[{}]", (int) rowIndex + 1, exceptionMsg));
                    titleReaderResult.rowIndex2msg.put((int) rowIndex, exceptionMsg);
                    
                    if (getParent().debugger) {
                        log.error("[DEBUGGER][HutoolTitleReader.handle] 数据处理异常: rowIndex={}, error={}", 
                            rowIndex, exceptionMsg, e);
                    }

                }
            }
        };
    }

    /**
     * 创建ExcelSaxReader
     *
     * @param handler RowHandler实例
     * @return ExcelSaxReader实例
     */
    ExcelSaxReader createSaxReader(RowHandler handler) throws IOException {
        boolean isXlsx = ExcelFileUtil.isXlsx(getParent().sourceFile.getInputStream());
        
        if (getParent().debugger) {
            log.info("[DEBUGGER][HutoolTitleReader.createSaxReader] 创建SaxReader: isXlsx={}, fileName={}", 
                isXlsx, getParent().sourceFile.getOriginalFilename());
        }
        
        return ExcelSaxUtil.createSaxReader(isXlsx, handler);
    }

    /**
     * 初始化SAX读取器（保留用于向后兼容）
     *
     * @return ExcelSaxReader实例
     */
    @Deprecated
    ExcelSaxReader initSaxReader() throws IOException {
        return createSaxReader(createRowHandler());
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
        super.TEMPLATE_CONSISTENCY_CHECK_ENABLE = false;
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
     * 禁用数据内存缓存
     * 适用于大数据量导入场景，可显著降低内存占用
     * 
     * @return HutoolTitleReader实例（支持链式调用）
     */
    public HutoolTitleReader<R> disableDataMemoryCache() {
        super.enableDataMemoryCache = false;
        return this;
    }


    @Override
    public Integer getSheetIndex() {
        // 优先从 Map 中查找（适用于通过 switchSheet 创建的 Reader）
        Integer index = getParent().sheetIndex2ReaderBiMap.getKey(this);
        if (index != null) {
            return index;
        }
        // 如果不在 Map 中（通过 reset 创建的 Reader），返回 sheetIndex 字段
        return this.sheetIndex;
    }

    @Override
    public TitleReaderResult<R> getReadResult() {
        return (TitleReaderResult) readerResult;
    }

}
