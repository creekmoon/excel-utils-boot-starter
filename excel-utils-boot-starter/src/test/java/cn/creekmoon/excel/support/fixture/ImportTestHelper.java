package cn.creekmoon.excel.support.fixture;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import org.springframework.mock.web.MockMultipartFile;

import java.io.IOException;
import java.io.InputStream;

public final class ImportTestHelper {

    private ImportTestHelper() {
    }

    public static MockMultipartFile loadFixture(String fileName) throws IOException {
        InputStream inputStream = ImportTestHelper.class.getResourceAsStream("/fixtures/import/" + fileName);
        if (inputStream == null) {
            throw new IllegalArgumentException("找不到 fixture: " + fileName);
        }
        return new MockMultipartFile(fileName, fileName, null, inputStream.readAllBytes());
    }

    public static ExcelImport createImport(String fileName) throws IOException {
        return ExcelImport.create(loadFixture(fileName));
    }

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
