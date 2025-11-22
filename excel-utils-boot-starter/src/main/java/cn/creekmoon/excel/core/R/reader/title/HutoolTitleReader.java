package cn.creekmoon.excel.core.R.reader.title;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.StringConverter;
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
        AtomicLong result = new AtomicLong(0);
        ExcelSaxReader<?> excel07SaxReader = ExcelSaxUtil.createSaxReader(ExcelFileUtil.isXlsx(getParent().sourceFile.getInputStream()), (new RowHandler() {
            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowCells) {
                if (sheetIndex != HutoolTitleReader.super.sheetIndex) {
                    return;
                }
                result.incrementAndGet();
            }
        }));
        try {
            excel07SaxReader.read(getParent().sourceFile.getInputStream(), -1);
        } catch (Exception e) {
            log.error("getSheetRowCount方法读取文件异常", e);
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
        /*尝试拿锁*/
        ExcelUtilsConfig.importParallelSemaphore.acquire();
        getReadResult().readStartTime = LocalDateTime.now();
        try {
            //新版读取 使用SAX读取模式

            ExcelSaxReader saxReader = initSaxReader();
            /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
            saxReader.read(this.getParent().sourceFile.getInputStream(), -1);
        } catch (Exception e) {
            log.error("SaxReader读取Excel文件异常", e);
        } finally {
            getReadResult().readSuccessTime = LocalDateTime.now();
            /*释放信号量*/
            ExcelUtilsConfig.importParallelSemaphore.release();
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
     * 初始化SAX读取器
     *
     * @param
     * @return
     */

    ExcelSaxReader initSaxReader() throws IOException {

        int targetSheetIndex = super.sheetIndex;
        TitleReaderResult titleReaderResult = (TitleReaderResult) getReadResult();

        /*返回一个Sax读取器*/
        return  ExcelSaxUtil.createSaxReader(ExcelFileUtil.isXlsx(getParent().sourceFile.getInputStream()),new RowHandler() {

            @Override
            public void doAfterAllAnalysed() {
                /*sheet读取结束时*/
            }


            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {
                if (targetSheetIndex != sheetIndex) {
                    return;
                }

                /*读取标题*/
                if (rowIndex == titleRowIndex) {
                    for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                        colIndex2Title.put(colIndex, StringConverter.parse(rowList.get(colIndex)));
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
                        return;
                    }
                    /*转换后置处理器*/
                    for (ExConsumer convertPostProcessor : convertPostProcessors) {
                        convertPostProcessor.accept(currentObject);
                    }
                    titleReaderResult.rowIndex2msg.put((int) rowIndex, CONVERT_SUCCESS_MSG);
                    /*消费*/
                    titleReaderResult.rowIndex2dataBiMap.put((int) rowIndex, currentObject);
                } catch (Exception e) {
                    //假如存在任一数据convert阶段就失败的单, 将打一个标记
                    titleReaderResult.EXISTS_READ_FAIL.set(true);
                    titleReaderResult.errorCount.incrementAndGet();
                    /*写入导出Excel结果*/
                    String exceptionMsg = GlobalExceptionMsgManager.getExceptionMsg(e);
                    getReadResult().errorReport.append(StrFormatter.format("第[{}]行发生错误[{}]", (int) rowIndex + 1, exceptionMsg));
                    titleReaderResult.rowIndex2msg.put((int) rowIndex, exceptionMsg);

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
