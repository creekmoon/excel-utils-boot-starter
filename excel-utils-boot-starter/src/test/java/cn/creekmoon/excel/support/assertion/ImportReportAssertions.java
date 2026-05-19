package cn.creekmoon.excel.support.assertion;

import cn.creekmoon.excel.core.R.report.ImportReport;
import lombok.extern.slf4j.Slf4j;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

@Slf4j
public final class ImportReportAssertions {

    private ImportReportAssertions() {
    }

    public static <T> void assertBalancedStatistics(ImportReport<T> report) {
        assertNotNull(report, "ImportReport should not be null");
        assertEquals(report.getTotalRows(), report.getSuccessRows() + report.getErrorRows(),
                "totalRows must equal successRows + errorRows");
        log.info("[ASSERT] report balanced: total={}, success={}, error={}",
                report.getTotalRows(), report.getSuccessRows(), report.getErrorRows());
    }

    public static <T> void assertNoErrors(ImportReport<T> report) {
        assertEquals(0, report.getErrorRows(), "errorRows should be 0");
        assertTrue(report.getGlobalErrors() == null || report.getGlobalErrors().isEmpty(),
                "globalErrors should be empty");
        log.info("[ASSERT] report has no errors, successRows={}", report.getSuccessRows());
    }

    public static <T> void logReportSummary(ImportReport<T> report) {
        log.info("[REPORT] sheet={}, total={}, success={}, error={}, globalErrors={}, rowErrorCount={}",
                report.getSheetName(),
                report.getTotalRows(),
                report.getSuccessRows(),
                report.getErrorRows(),
                report.getGlobalErrors(),
                report.getRowErrors() == null ? 0 : report.getRowErrors().size());
    }
}
