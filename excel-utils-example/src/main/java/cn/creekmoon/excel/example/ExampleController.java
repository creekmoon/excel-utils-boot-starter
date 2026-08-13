package cn.creekmoon.excel.example;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.report.ImportReport;
import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.core.W.title.TitleWriter;
import cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.ResponseEntity;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle.of;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.*;
import static cn.creekmoon.excel.example.Student.createStudentList;
import static cn.creekmoon.excel.example.Teacher.createTeacherList;

/**
 * Excel 导入导出用法示例
 *
 * <p>按业务场景分类，每个接口聚焦一种能力，避免重复。</p>
 *
 * <p><b>导出场景（Export）</b></p>
 * <ul>
 *   <li>基础导出（单/多 Sheet）</li>
 *   <li>条件样式与自定义格式</li>
 *   <li>多级表头</li>
 *   <li>范围控制（range）</li>
 *   <li>同一 Sheet 多表格（reset）</li>
 *   <li>大数据量导出</li>
 * </ul>
 *
 * <p><b>导入场景（Import）</b></p>
 * <ul>
 *   <li>基础导入（含转换器）</li>
 *   <li>返回 JSON 报告（API 场景）</li>
 *   <li>自定义配置（fail-fast / 错误阈值 / 必填校验）</li>
 *   <li>范围控制导入</li>
 *   <li>指定单元格读取</li>
 *   <li>同一 Sheet 多表格（reset）</li>
 *   <li>后置处理器（跨字段校验）</li>
 * </ul>
 */
@Slf4j
@Tag(name = "Excel 导入导出示例")
@RestController
public class ExampleController {

    /* ========================= 公共工具方法 ========================= */

    private MultipartFile getTestFile(String resourceName) throws IOException {
        return new MockMultipartFile(
                resourceName,
                resourceName,
                "application/vnd.ms-excel",
                new ClassPathResource(resourceName).getInputStream()
        );
    }

    private MultipartFile fallbackFile(MultipartFile file, String resourceName) throws IOException {
        return file != null ? file : getTestFile(resourceName);
    }


    /* ================================================================
     *                        导出示例（Export）
     * ================================================================ */

    @GetMapping("/export/basic")
    @Operation(summary = "基础导出", description = "单 Sheet + 多 Sheet 基础导出")
    public void exportBasic(Integer size, HttpServletResponse response) throws IOException {
        ArrayList<Student> students = createStudentList(size != null ? size : 10_000);
        ArrayList<Teacher> teachers = createTeacherList(size != null ? size / 2 : 5_000);

        ExcelExport excelExport = ExcelExport.create();

        // Sheet1: 学生信息（多级表头）
        excelExport.switchNewSheet("学生信息", Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName)
                .addTitle("联系方式::邮箱", Student::getEmail)
                .addTitle("其他::年龄", Student::getAge)
                .addTitle("其他::生日", Student::getBirthday)
                .addTitle("其他::过期时间", Student::getExpTime)
                .write(students);

        // Sheet2: 教师信息
        excelExport.switchNewSheet("教师信息", Teacher.class)
                .addTitle("教师姓名", Teacher::getTeacherName)
                .addTitle("工作年限", Teacher::getWorkYear)
                .write(teachers);

        excelExport.response(response);
    }

    @GetMapping("/export/with-style")
    @Operation(summary = "条件样式与自定义格式", description = "根据数据条件动态着色 + 自定义数字格式")
    public void exportWithStyle(HttpServletResponse response) throws IOException {
        ArrayList<Student> students = createStudentList(1_000);

        // 自定义数字格式：保留两位小数
        ExcelCellStyle twoDecimalStyle = new ExcelCellStyle(
                (workbook, style) -> style.setDataFormat(
                        workbook.createDataFormat().getFormat("#,##0.00")
                )
        );

        ExcelExport excelExport = ExcelExport.create();
        excelExport.switchNewSheet(Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName,
                        of(LIGHT_GREEN, s -> s.getAge() > 50))
                .addTitle("年龄", Student::getAge,
                        of(twoDecimalStyle, s -> s.getAge() > 50))
                .addTitle("邮箱", Student::getEmail)
                .write(students);

        excelExport.response(response);
    }

    @GetMapping("/export/with-range")
    @Operation(summary = "范围控制导出", description = "使用 range 控制标题和数据写入位置")
    public void exportWithRange(HttpServletResponse response) throws IOException {
        ArrayList<Student> students = createStudentList(100);

        ExcelExport excelExport = ExcelExport.create();

        // 示例1: 从第6行开始，限制写到第20行
        excelExport.switchNewSheet("范围示例1", Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .range(5, 19)
                .write(students);

        // 示例2: 从第10行开始，不限制结束行
        excelExport.switchNewSheet("范围示例2", Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge,
                        of(LIGHT_GREEN, s -> s.getAge() > 50))
                .range(9)
                .write(students);

        // 示例3: 多级表头 + 完整指定范围
        excelExport.switchNewSheet("多级表头+范围", Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName,
                        of(LIGHT_ORANGE, s -> true))
                .addTitle("年龄", Student::getAge)
                .range(2, 4, 15)
                .write(students);

        excelExport.response(response);
    }

    @GetMapping("/export/with-reset")
    @Operation(summary = "同一 Sheet 多表格", description = "使用 reset 在同一 Sheet 中写入不同类型数据")
    public void exportWithReset(HttpServletResponse response) throws IOException {
        ArrayList<Student> students = createStudentList(50);
        ArrayList<Teacher> teachers = createTeacherList(30);

        ExcelExport excelExport = ExcelExport.create();
        TitleWriter<Student> writer = excelExport.switchNewSheet("多表格Sheet", Student.class);

        writer.addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .write(students);

        writer.reset(Teacher.class)
                .addTitle("教师姓名", Teacher::getTeacherName)
                .addTitle("工作年限", Teacher::getWorkYear,
                        of(LIGHT_GREEN, t -> t.getWorkYear() > 5))
                .write(teachers);

        writer.reset(Student.class, 2)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .write(students.subList(0, 10));

        excelExport.response(response);
    }

    @GetMapping("/export/large-data")
    @Operation(summary = "大数据量导出", description = "导出 10 万行数据测试性能")
    public void exportLargeData(HttpServletResponse response) throws IOException {
        ArrayList<Student> students = createStudentList(100_000);

        long start = System.currentTimeMillis();
        ExcelExport excelExport = ExcelExport.create();
        excelExport.switchNewSheet(Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(students);

        excelExport.response(response);
        log.info("大数据导出完成，耗时 {} ms", System.currentTimeMillis() - start);
    }

    @GetMapping("/export/style-showcase")
    @Operation(summary = "预设样式展示", description = "输出所有内置预设样式，供开发者参考")
    public void exportStyleShowcase(HttpServletResponse response) throws IOException {
        List<StyleDemo> data = Arrays.asList(
                StyleDemo.create("张三", 25, new BigDecimal("12345.67"), 0.15,
                        LocalDate.now(), LocalDateTime.now(), "正常", "A类", 85),
                StyleDemo.create("李四", 45, new BigDecimal("99999.99"), 0.85,
                        LocalDate.now(), LocalDateTime.now(), "警告", "B类", 72),
                StyleDemo.create("王五", 65, new BigDecimal("500.00"), 0.05,
                        LocalDate.now(), LocalDateTime.now(), "异常", "C类", 58),
                StyleDemo.create("赵六", 35, new BigDecimal("8888.88"), 0.50,
                        LocalDate.now(), LocalDateTime.now(), "正常", "A类", 91)
        );

        ExcelExport excelExport = ExcelExport.create();

        /* ===== Sheet 1: 表头样式 ===== */
        excelExport.switchNewSheet("1-表头样式", StyleDemo.class)
                .addTitle("姓名", StyleDemo::getName)
                .addTitle("年龄", StyleDemo::getAge)
                .addTitle("分数", StyleDemo::getScore)
                .write(data);

        /* ===== Sheet 2: 状态色（条件样式） ===== */
        excelExport.switchNewSheet("2-状态色", StyleDemo.class)
                .addTitle("姓名", StyleDemo::getName)
                .addTitle("状态", StyleDemo::getStatus,
                        of(STATUS_SUCCESS, d -> "正常".equals(d.getStatus())),
                        of(STATUS_WARNING, d -> "警告".equals(d.getStatus())),
                        of(STATUS_DANGER, d -> "异常".equals(d.getStatus())))
                .addTitle("分数", StyleDemo::getScore,
                        of(STATUS_SUCCESS_BORDER, d -> d.getScore() >= 80),
                        of(STATUS_WARNING_BORDER, d -> d.getScore() >= 60 && d.getScore() < 80),
                        of(STATUS_DANGER_BORDER, d -> d.getScore() < 60))
                .write(data);

        /* ===== Sheet 3: 斑马纹 + 边框 ===== */
        excelExport.switchNewSheet("3-斑马纹+边框", StyleDemo.class)
                .addTitle("姓名", StyleDemo::getName, BORDER_THIN_ALL, d -> true)
                .addTitle("年龄", StyleDemo::getAge, BORDER_THIN_ALL, d -> true)
                .addTitle("类别", StyleDemo::getCategory, BORDER_THIN_ALL, d -> true)
                .addTitle("分数", StyleDemo::getScore, BORDER_THIN_ALL, d -> true)
                .write(data);

        /* ===== Sheet 4: 数字格式 ===== */
        excelExport.switchNewSheet("4-数字格式", StyleDemo.class)
                .addTitle("姓名", StyleDemo::getName)
                .addTitle("货币", StyleDemo::getAmount, FORMAT_CURRENCY, d -> true)
                .addTitle("百分比", StyleDemo::getRate, FORMAT_PERCENTAGE, d -> true)
                .addTitle("千分位", StyleDemo::getScore, FORMAT_INTEGER, d -> true)
                .addTitle("日期", StyleDemo::getDate, FORMAT_DATE, d -> true)
                .addTitle("日期时间", StyleDemo::getDatetime, FORMAT_DATETIME, d -> true)
                .write(data);

        /* ===== Sheet 5: 对齐样式 ===== */
        excelExport.switchNewSheet("5-对齐样式", StyleDemo.class)
                .addTitle("左对齐", StyleDemo::getName, ALIGN_LEFT, d -> true)
                .addTitle("居中", StyleDemo::getStatus, ALIGN_CENTER, d -> true)
                .addTitle("右对齐", StyleDemo::getScore, ALIGN_RIGHT, d -> true)
                .write(data);

        /* ===== Sheet 6: 组合样式 ===== */
        excelExport.switchNewSheet("6-组合样式", StyleDemo.class)
                .addTitle("表头+边框+居中", StyleDemo::getName,
                        ExcelCellStyle.merge(HEADER_PRIMARY, BORDER_THIN_ALL, ALIGN_CENTER), d -> true)
                .addTitle("状态+边框+右对齐", StyleDemo::getScore,
                        of(ExcelCellStyle.merge(STATUS_SUCCESS, BORDER_THIN_ALL, ALIGN_RIGHT),
                                d -> d.getScore() >= 80))
                .write(data);

        /* ===== Sheet 7: 自定义颜色 ===== */
        excelExport.switchNewSheet("7-自定义颜色", StyleDemo.class)
                .addTitle("HEX 色值", StyleDemo::getName,
                        of(ExcelCellStyle.hex(0xE8D5B7), d -> true))
                .addTitle("RGB 色值", StyleDemo::getStatus,
                        of(ExcelCellStyle.rgb(180, 220, 250), d -> true))
                .addTitle("IndexedColor", StyleDemo::getCategory,
                        of(ExcelCellStyle.hex(0xDDA0DD), d -> true))
                .write(data);

        /* ===== Sheet 8: 字体样式 ===== */
        excelExport.switchNewSheet("8-字体样式", StyleDemo.class)
                .addTitle("粗体", StyleDemo::getName, FONT_BOLD, d -> true)
                .addTitle("斜体", StyleDemo::getStatus, FONT_ITALIC, d -> true)
                .addTitle("粗体+居中", StyleDemo::getScore,
                        ExcelCellStyle.merge(FONT_BOLD, ALIGN_CENTER), d -> true)
                .write(data);

        excelExport.response(response);
    }


    /* ================================================================
     *                        导入示例（Import）
     * ================================================================ */

    @PostMapping("/import/basic")
    @Operation(summary = "基础导入", description = "基础按行读取 + 类型转换 + 流式消费")
    public void importBasic(MultipartFile file, HttpServletResponse response) throws Exception {
        file = fallbackFile(file, "导入测试.xlsx");

        ExcelImport excelImport = ExcelImport.create(file);

        // Sheet1: 基础导入
        excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .read(s -> log.info("Sheet1 读取: {}", s));

        // Sheet2: 只读取部分列
        excelImport.switchSheet(1, Student::new)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("邮箱", Student::setEmail)
                .read(s -> log.info("Sheet2 读取: {}", s));

        excelImport.response(response);
    }

    @PostMapping("/import/with-report")
    @Operation(summary = "导入并返回 JSON 报告", description = "API 场景：不生成 Excel，直接返回结构化报告")
    public ResponseEntity<ImportReport> importWithReport(MultipartFile file) throws Exception {
        file = fallbackFile(file, "导入测试.xlsx");

        ImportReport report = ExcelImport.create(file)
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .readWithReport();

        log.info("导入完成：总计 {} 条，成功 {} 条，失败 {} 条",
                report.getTotalRows(), report.getSuccessRows(), report.getErrorRows());

        return ResponseEntity.ok(report);
    }

    @PostMapping("/import/report-and-excel")
    @Operation(summary = "一次读取同时产出报告与结果文件", description = "同一次读取的结果既可取 JSON 报告，也可回写错误列到 Excel")
    public void importReportAndExcel(MultipartFile file, HttpServletResponse response) throws Exception {
        file = fallbackFile(file, "导入测试.xlsx");

        ExcelImport excelImport = ExcelImport.create(file);
        ImportReport report = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .readWithReport();

        log.info("导入完成：总计 {} 条，成功 {} 条，失败 {} 条",
                report.getTotalRows(), report.getSuccessRows(), report.getErrorRows());

        if (report.hasGlobalErrors()) {
            log.warn("全局错误：{} - {}", report.getGlobalErrors(), report.getGlobalErrorDisplayMessage());
        }

        excelImport.response(response);
    }

    @PostMapping("/import/with-range")
    @Operation(summary = "范围控制导入", description = "只读取指定范围内的数据")
    public void importWithRange(MultipartFile file, HttpServletResponse response) throws Exception {
        file = fallbackFile(file, "导入测试.xlsx");

        ExcelImport excelImport = ExcelImport.create(file);

        // 只读取前 20 行
        excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .range(0, 19)
                .read(s -> log.info("Sheet1: {}", s));

        // 从第 5 行开始读取
        excelImport.switchSheet(1, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .range(4)
                .read(s -> log.info("Sheet2: {}", s));

        excelImport.response(response);
    }

    @PostMapping("/import/with-cell-reader")
    @Operation(summary = "指定单元格读取", description = "读取特定单元格而非按行遍历")
    public void importWithCellReader(MultipartFile file, HttpServletResponse response) throws Exception {
        file = fallbackFile(file, "导入测试(读取指定单元格).xlsx");

        Student student = ExcelImport.create(file)
                .switchSheetAndUseCellReader(1, Student::new)
                .addConvertAndMustExist("A1", Student::setUserName)
                .read(s -> log.info("读取到: {}", s));

        log.info("单元格读取结果: {}", student);

        // 结果回写（无错误数据，仅演示 response 用法）
        ExcelImport.create(file).response(response);
    }

    @PostMapping("/import/with-reset")
    @Operation(summary = "同一 Sheet 多表格导入", description = "使用 reset 读取同一 Sheet 中不同类型数据")
    public void importWithReset(HttpServletResponse response) throws Exception {
        MultipartFile file = getTestFile("导入测试Sheet页存在多个表格.xlsx");

        ExcelImport excelImport = ExcelImport.create(file);

        // 第一个表格：Student（第 0 行开始）
        TitleReader<Student> reader = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .range(0, 50);
        reader.read(s -> log.info("Student: {}", s));

        // 第二个表格：Teacher（约第 52 行）
        TitleReader<Teacher> teacherReader = reader.reset(Teacher::new)
                .addConvert("教师姓名", Teacher::setTeacherName)
                .addConvert("工作年限", IntegerConverter::parse, Teacher::setWorkYear)
                .range(52, 82);
        teacherReader.read(t -> log.info("Teacher: {}", t));

        // 第三个表格：Student（约第 84 行）
        teacherReader.reset(Student::new)
                .addConvert("基本信息::用户名", Student::setUserName)
                .addConvert("基本信息::全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .range(84)
                .read(s -> log.info("Student(多级表头): {}", s));

        // reset 创建的 Reader 不参与结果文件生成，不调用 response
        log.info("多表格读取完成");
    }

    @PostMapping("/import/with-post-processor")
    @Operation(summary = "后置处理器", description = "convert 完成后执行跨字段校验")
    public ResponseEntity<ImportReport> importWithPostProcessor(MultipartFile file) throws Exception {
        file = fallbackFile(file, "导入测试.xlsx");

        ImportReport report = ExcelImport.create(file)
                .switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvertPostProcessor(student -> {
                    // 跨字段校验：如果年龄大于 60，邮箱必须以 @qq.com 结尾
                    if (student.getAge() != null && student.getAge() > 60) {
                        if (student.getEmail() == null || !student.getEmail().endsWith("@qq.com")) {
                            throw new RuntimeException("年龄大于 60 的用户邮箱必须以 @qq.com 结尾");
                        }
                    }
                })
                .readWithReport();

        return ResponseEntity.ok(report);
    }

    @PostMapping("/import/with-validation")
    @Operation(summary = "字段级校验", description = "必填字段 + 可空字段 + 跳过空白")
    public ResponseEntity<ImportReport> importWithValidation(MultipartFile file) throws Exception {
        file = fallbackFile(file, "导入测试.xlsx");

        ImportReport report = ExcelImport.create(file)
                .switchSheet(0, Student::new)
                .addConvertAndMustExist("用户名", Student::setUserName)       // 必填
                .addConvertAndMustExist("全名", Student::setFullName)         // 必填
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)  // 普通转换
                .addConvertAndSkipEmpty("邮箱", Student::setEmail)             // 空白则跳过
                .readWithReport();

        return ResponseEntity.ok(report);
    }
}
