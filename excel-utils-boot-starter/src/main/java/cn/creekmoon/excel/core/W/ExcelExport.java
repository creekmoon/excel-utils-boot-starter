package cn.creekmoon.excel.core.W;


import cn.creekmoon.excel.core.W.title.HutoolTitleWriter;
import cn.creekmoon.excel.core.W.title.TitleWriter;
import cn.creekmoon.excel.util.ExcelFileUtils;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.lang.UUID;
import cn.hutool.core.map.BiMap;
import cn.hutool.core.text.StrFormatter;
import jakarta.servlet.http.HttpServletResponse;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;
import java.util.HashMap;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelExport {
    /**
     * 打印调试内容
     */
    public boolean debugger = false;

    /*唯一识别名称*/
    @Getter
    protected String taskId = UUID.randomUUID().toString();
    /*生成的临时文件路径*/
    @Getter
    protected String resultFilePath = ExcelFileUtils.generateXlsxAbsoluteFilePath(taskId);

    /*自定义的名称*/
    public String excelName;

    /*各个写入器的共享缓存, 自己想存点啥就放里面*/
    public HashMap<Object, Object> metadatas = new HashMap<>();


    /*
     * sheet页和导出对象的映射关系
     * */
    public BiMap<Integer, Writer> sheetIndex2SheetWriter = new BiMap<>(new HashMap<>());


    private ExcelExport() {

    }


    public static ExcelExport create() {
        ExcelExport excelExport = new ExcelExport();
        excelExport.excelName = "export_result_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy_MM_dd_HH_mm")) + ".xlsx";
        return excelExport;
    }


    /**
     * 异步创建excel
     *
     * @param async
     * @param successHandler 成功回调<taskId, 文件>
     * @return taskId
     */
//    public static String createAsync(Consumer<ExcelExport> async, BiConsumer<String, File> successHandler) {
//        ExcelExport excelExport = create();
//        Thread.ofVirtual().start(() -> {
//            async.accept(excelExport);
//            File tile = excelExport.stopWrite();
//            ExcelFileUtils.cleanTempFileByPathDelay(tile);
//            successHandler.accept(excelExport.taskId, tile);
//        });
//        return excelExport.taskId;
//    }


    /**
     * 切换到新的标签页
     */
    public <T> TitleWriter<T> switchSheet(Integer sheetIndex, Class<T> newDataClass) {
        if (sheetIndex2SheetWriter.containsKey(sheetIndex)) {
            return (TitleWriter<T>) sheetIndex2SheetWriter.get(sheetIndex);
        }
        if (!sheetIndex2SheetWriter.containsKey(sheetIndex - 1)) {
            throw new RuntimeException(StrFormatter.format("切换sheet页失败,预期切换到sheet索引={}, 但当前最大索引下标sheet={},请检查代码并保证sheet页下标连续!"));
        }
        return switchNewSheet(newDataClass);
    }


    /**
     * 切换到新的标签页
     */
    public <T> TitleWriter<T> switchNewSheet(Class<T> newDataClass) {
        HutoolTitleWriter<T> newTitleWriter = new HutoolTitleWriter<T>(this, sheetIndex2SheetWriter.size());
        sheetIndex2SheetWriter.put(newTitleWriter.getSheetIndex(), newTitleWriter);
        return newTitleWriter;
    }

    public ExcelExport debug() {
        this.debugger = true;
        return this;
    }


    /**
     * 响应并清除文件
     *
     * @param response
     * @throws IOException
     */
    public void response(HttpServletResponse response) throws IOException {
        File file = this.stopWrite();
        ExcelFileUtils.response(file, excelName, response);
    }


    /**
     * 停止写入
     * 1.进行手动重排序  因为在wps/excel中sheet显示顺序与index无关，还有隐藏sheet
     * 2. 回调各个写入器的onStopWrite方法
     *
     * @return 结果文件绝对路径
     */
    public File stopWrite() {

        Workbook workbook = sheetIndex2SheetWriter.get(0).getWorkbook();
        sheetIndex2SheetWriter.values()
                .stream()
                .sorted(Comparator.comparing(Writer::getSheetIndex))
                .peek(x -> workbook.setSheetOrder(x.getSheetName(), x.getSheetIndex()))
                .forEach(Writer::stopWrite);

        return new File(getResultFilePath());
    }


}
