package cn.creekmoon.excel.util;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** ExcelFileUtils 临时文件路径与清理契约。 */
class ExcelFileUtilsTest {

    @TempDir
    Path tempDir;

    @Test
    void generateXlsxAbsoluteFilePath_endsWithTempFilesAndXlsx() {
        String path = ExcelFileUtils.generateXlsxAbsoluteFilePath("unique-id-123");
        assertTrue(path.endsWith("temp_files" + File.separator + "unique-id-123.xlsx"), path);
    }

    @Test
    void generateXlsxAbsoluteFilePath_createsParentDirectory() {
        String path = ExcelFileUtils.generateXlsxAbsoluteFilePath("unique-id-123");
        assertTrue(Files.isDirectory(Path.of(path).getParent()), "temp_files dir must exist");
    }

    @Test
    void getCustomAbsoluteFilePath_appendsFileNameUnderTempFiles() {
        String path = ExcelFileUtils.getCustomAbsoluteFilePath("report.csv");
        assertTrue(path.endsWith("temp_files" + File.separator + "report.csv"), path);
    }

    @Test
    void cleanTempFileByPathNow_deletesExistingFile() throws IOException {
        Path file = tempDir.resolve("to-clean.xlsx");
        Files.write(file, new byte[]{1, 2, 3});
        assertTrue(Files.exists(file));

        ExcelFileUtils.cleanTempFileByPathNow(file.toString());
        assertFalse(Files.exists(file), "file should be deleted");
    }

    @Test
    void cleanTempFileByPathNow_missingFile_doesNotThrow() {
        ExcelFileUtils.cleanTempFileByPathNow(tempDir.resolve("never-existed.xlsx").toString());
    }
}
