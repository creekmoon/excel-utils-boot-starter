package cn.creekmoon.excel.support.fixture;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import org.springframework.mock.web.MockMultipartFile;

import java.io.IOException;
import java.io.InputStream;

/** 导入测试公共入口：加载 fixture、绑定标准列映射。 */
public final class ImportTestHelper {

    private ImportTestHelper() {
    }

    /** 从 classpath:/fixtures/import/ 加载为 MockMultipartFile。 */
    public static MockMultipartFile loadFixture(String fileName) throws IOException {
        InputStream inputStream = ImportTestHelper.class.getResourceAsStream("/fixtures/import/" + fileName);
        if (inputStream == null) {
            throw new IllegalArgumentException("找不到 fixture: " + fileName);
        }
        return new MockMultipartFile(fileName, fileName, null, inputStream.readAllBytes());
    }

    /** 创建 ExcelImport 并绑定指定 fixture 文件。 */
    public static ExcelImport createImport(String fileName) throws IOException {
        return ExcelImport.create(loadFixture(fileName));
    }

    /** 标准学生表列：用户名、全名、年龄、邮箱、生日、过期时间。 */
    public static TitleReader<Student> bindStandardColumns(TitleReader<Student> reader) {
        return reader
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime);
    }
}
