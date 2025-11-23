package cn.creekmoon.excel.core.R;

import cn.creekmoon.excel.core.R.reader.Reader;
import cn.creekmoon.excel.core.R.reader.cell.CellReader;
import cn.creekmoon.excel.core.R.reader.cell.HutoolCellReader;
import cn.creekmoon.excel.core.R.reader.title.HutoolTitleReader;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.ExcelFileUtils;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.lang.UUID;
import cn.hutool.core.map.BiMap;
import cn.hutool.core.text.csv.CsvReader;
import cn.hutool.core.text.csv.CsvRow;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;
import jakarta.servlet.http.HttpServletResponse;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.function.Supplier;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelImport {

    public ArrayList<Reader> allReaders = new ArrayList<>();

    /**
     * Sheet的rId和SheetName的双向映射
     * 该映射基于Excel文件的workbook.xml解析，保证有序且稳定
     */
    public BiMap<String, String> rid2SheetNameBiMap;

    /**
     * 打印调试内容
     */
    public boolean debugger = false;
    /*唯一识别名称 会同步生成一份文件到临时目录*/
    public String taskId = UUID.fastUUID().toString();
    String resultFilePath = ExcelFileUtils.generateXlsxAbsoluteFilePath(taskId);

    /*当前导入的文件*/
    public MultipartFile sourceFile;


    private ExcelImport() {
    }


    public static ExcelImport create(MultipartFile file) throws IOException {
        return create(file, false);
    }


    public static ExcelImport create(MultipartFile file, boolean debugger) throws IOException {
        ExcelImport excelImport = new ExcelImport();
        excelImport.sourceFile = file;
        excelImport.debugger = debugger;
        excelImport.csvSupport();
        // 解析Excel文件的sheet映射关系（rId ↔ SheetName）
        excelImport.rid2SheetNameBiMap = excelImport.parseSheetMappings();

        return excelImport;
    }


    /**
     * 切换读取的sheet页
     *
     * @param sheetIndex 下标,从0开始
     * @param supplier   按行读取时,每行数据的实例化对象构造函数
     * @param <T>
     * @return
     */
    public <T> TitleReader<T> switchSheet(int sheetIndex, Supplier<T> supplier) {
        if (supplier.get() == supplier.get()) {
            throw new RuntimeException("导出的对象不支持重写 equal() 方法. 请勿使用@Data注解!");
        }

        // 从有序BiMap中获取第sheetIndex个entry的rId和sheetName
        List<String> ridList = new ArrayList<>(rid2SheetNameBiMap.keySet());
        if (debugger) {
            log.info("[DEBUGGER][ExcelImport.switchSheet] 请求sheetIndex={}, 可用的rId列表: {}", sheetIndex, ridList);
        }
        if (sheetIndex < 0 || sheetIndex >= ridList.size()) {
            throw new RuntimeException("Sheet index out of bounds: " + sheetIndex + ", available sheets: " + ridList.size());
        }
        String rId = ridList.get(sheetIndex);
        String sheetName = rid2SheetNameBiMap.get(rId);

        if (debugger) {
            log.info("[DEBUGGER][ExcelImport.switchSheet] 选择的sheet: index={}, rId={}, sheetName={}",
                    sheetIndex, rId, sheetName);
        }

        //新增读取器，传递rId和sheetName
        HutoolTitleReader<T> reader = new HutoolTitleReader<>(this, rId, sheetName, supplier);
        this.allReaders.add(reader);
        return reader;
    }


    /**
     * 切换读取的sheet页
     *
     * @param sheetIndex 下标,从0开始
     * @param supplier   按行读取时,每行数据的实例化对象构造函数
     * @param <T>
     * @return
     */
    public <T> CellReader<T> switchSheetAndUseCellReader(int sheetIndex, Supplier<T> supplier) {

        // 从有序BiMap中获取第sheetIndex个entry的rId和sheetName
        List<String> ridList = new ArrayList<>(rid2SheetNameBiMap.keySet());
        if (debugger) {
            log.info("[DEBUGGER][ExcelImport.switchSheetAndUseCellReader] 请求sheetIndex={}, 可用的rId列表: {}", sheetIndex, ridList);
        }
        if (sheetIndex < 0 || sheetIndex >= ridList.size()) {
            throw new RuntimeException("Sheet index out of bounds: " + sheetIndex + ", available sheets: " + ridList.size());
        }
        String rId = ridList.get(sheetIndex);
        String sheetName = rid2SheetNameBiMap.get(rId);

        if (debugger) {
            log.info("[DEBUGGER][ExcelImport.switchSheetAndUseCellReader] 选择的sheet: index={}, rId={}, sheetName={}",
                    sheetIndex, rId, sheetName);
        }

        //新增读取器，传递rId和sheetName
        HutoolCellReader<T> reader = new HutoolCellReader<>(this, rId, sheetName, supplier);
        this.allReaders.add(reader);
        return reader;
    }


    /**
     * 支持csv类型的文件  本质是内部将csv转为xlsx
     *
     * @return
     */
    @SneakyThrows
    protected ExcelImport csvSupport() {

        if (StrUtil.isBlank(this.sourceFile.getOriginalFilename()) || !this.sourceFile.getOriginalFilename().toLowerCase().contains(".csv")) {
            /*如果不是csv文件 跳过这个方法*/
            return this;
        }
        log.info("[文件导入]收到CSV格式的文件[{}],尝试转化为XLSX格式", this.sourceFile.getOriginalFilename());
        /*获取CSV文件并尝试读取*/
        String filePath = ExcelFileUtils.generateXlsxAbsoluteFilePath(UUID.fastUUID().toString(true));
        BigExcelWriter bigWriter = ExcelUtil.getBigWriter();
        try {
            CsvReader read = new CsvReader(new InputStreamReader(sourceFile.getInputStream()), null);
            Iterator<CsvRow> rowIterator = read.iterator();
            while (rowIterator.hasNext()) {
                bigWriter.writeRow(rowIterator.next().getRawList());
            }
        } catch (Exception e) {
            log.error("[文件导入]csv转换异常!", e);
            this.sourceFile = null;
        } finally {
            bigWriter.close();
            ExcelFileUtils.cleanTempFileByPathDelay(filePath, 120);
        }

        /*将新的xlsx文件替换为当前的文件*/
        this.sourceFile = new MockMultipartFile("csv2xlsx.xlsx", FileUtil.getInputStream(filePath));
        return this;
    }


    public ExcelImport response(HttpServletResponse response) throws IOException {
        File file = generateResultFile();
//        IoUtil.copy(sourceFile.getInputStream(), FileUtil.getOutputStream(ExcelFileUtils.getAbsoluteFilePath(this.taskId)));
//        File file = FileUtil.file(ExcelFileUtils.getAbsoluteFilePath(this.taskId));
//        System.out.println("file.canRead() = " + file.canRead());
//        System.out.println("file.canWrite() = " + file.canWrite());
        ExcelFileUtils.response(ExcelFileUtils.generateXlsxAbsoluteFilePath(this.taskId), taskId + ".xlsx", response);
        return this;
    }

    /**
     * 生成导入结果
     *
     * @return taskId
     */
    public File generateResultFile() throws IOException {
        return this.generateResultFile(true);
    }


    /**
     * 生成导入结果
     *
     * @param autoClean 是否自动删除临时文件(后台进行延迟删除)
     * @return File  生成的新结果文件
     */
    public File generateResultFile(boolean autoClean) throws IOException {

        String absoluteFilePath = resultFilePath;

        try (Workbook workbook = new XSSFWorkbook(sourceFile.getInputStream());
             BufferedOutputStream outputStream = FileUtil.getOutputStream(absoluteFilePath)) {
            for (Reader<?> reader : allReaders) {
                Sheet sheet = workbook.getSheet(reader.sheetName);
                if (reader instanceof TitleReader<?> titleReader) {

                    // 检查 colIndex2Title 是否已初始化（即是否调用过 read()）
                    if (titleReader.colIndex2Title == null || titleReader.colIndex2Title.isEmpty()) {
                        // Reader 未执行 read() 或 Sheet 不存在，跳过此 Reader
                        continue;
                    }

                    // 推算准备要写的位置
                    int titleRowIndex = titleReader.titleRowIndex;
                    Integer lastTitleColumnIndex = titleReader.colIndex2Title.keySet().stream().max(Integer::compareTo).get();
                    int msgTitleColumnIndex = lastTitleColumnIndex + 1;
                    Integer dataFirstRowIndex = titleReader.rowIndex2msg.keySet().stream().min(Integer::compareTo).orElse(null);
                    Integer dataLatestRowIndex = titleReader.rowIndex2msg.keySet().stream().max(Integer::compareTo).orElse(null);

                    // 添加空值检查
                    if (dataFirstRowIndex == null || dataLatestRowIndex == null) {
                        continue;
                    }

                    // 开始写结果行
                    CellStyle titleCellStyle = sheet.getRow(titleRowIndex).getCell(lastTitleColumnIndex).getCellStyle();
                    Cell cell1 = sheet.getRow(titleRowIndex).createCell(msgTitleColumnIndex);
                    cell1.setCellStyle(titleCellStyle);
                    cell1.setCellValue(ExcelConstants.RESULT_TITLE);

                    // 设置导入结果内容
                    for (Integer rowIndex = dataFirstRowIndex; rowIndex <= dataLatestRowIndex; rowIndex++) {
                        Cell cell = sheet.getRow(rowIndex).createCell(msgTitleColumnIndex);
                        cell.setCellValue(titleReader.rowIndex2msg.get(rowIndex));
                    }
                }
            }
            workbook.write(outputStream);
            outputStream.flush();
        }
        return FileUtil.file(absoluteFilePath);
    }

    /**
     * 解析Excel文件的workbook.xml，构建rId ↔ SheetName的双向映射
     * 使用POI Package API进行轻量级解析，不加载完整workbook
     *
     * @return rId ↔ SheetName 的双向映射表（有序）
     */
    private BiMap<String, String> parseSheetMappings() {
        BiMap<String, String> result = new BiMap<>(new LinkedHashMap<>());

        try {
            if (debugger) {
                log.info("[DEBUGGER][ExcelImport.parseSheetMappings] 开始解析workbook.xml");
            }

            // 使用POI Package API打开Excel文件（ZIP格式）
            try (OPCPackage pkg = OPCPackage.open(sourceFile.getInputStream())) {
                // 获取workbook part
                PackageRelationship workbookRel = pkg.getRelationshipsByType(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                ).getRelationship(0);

                PackagePart workbookPart = pkg.getPart(workbookRel);

                // 解析workbook.xml
                try (InputStream workbookStream = workbookPart.getInputStream()) {
                    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                    factory.setNamespaceAware(true);
                    DocumentBuilder builder = factory.newDocumentBuilder();
                    Document doc = builder.parse(workbookStream);

                    // 获取所有sheet节点（按XML顺序即为显示顺序）
                    NodeList sheetNodes = doc.getElementsByTagName("sheet");

                    for (int i = 0; i < sheetNodes.getLength(); i++) {
                        Element sheetElement = (Element) sheetNodes.item(i);

                        // 获取r:id属性（可能在不同命名空间）
                        String rid = sheetElement.getAttribute("r:id");
                        if (rid == null || rid.isEmpty()) {
                            rid = sheetElement.getAttributeNS(
                                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                                    "id"
                            );
                        }

                        String sheetName = sheetElement.getAttribute("name");

                        if (rid != null && !rid.isEmpty() && sheetName != null && !sheetName.isEmpty()) {
                            result.put(rid, sheetName);

                            if (debugger) {
                                log.info("[DEBUGGER][ExcelImport.parseSheetMappings] 解析Sheet: index={}, name={}, rid={}",
                                        i, sheetName, rid);
                            }
                        }
                    }
                }
            }

            if (debugger) {
                log.info("[DEBUGGER][ExcelImport.parseSheetMappings] 解析完成: 共{}个sheet", result.size());
            }

            if (result.isEmpty()) {
                throw new RuntimeException("[ExcelImport.parseSheetMappings] 无法解析Excel文件的Sheet信息，文件可能损坏");
            }

        } catch (Exception e) {
            log.error("[ExcelImport.parseSheetMappings] 解析workbook.xml失败", e);
            if (debugger) {
                log.error("[DEBUGGER][ExcelImport.parseSheetMappings] 解析异常详情", e);
            }
            throw new RuntimeException("[ExcelImport.parseSheetMappings] 解析Excel文件失败: " + e.getMessage(), e);
        }

        return result;
    }
}
