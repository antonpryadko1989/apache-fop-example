package com.visoft.services;

import com.visoft.dto.SummaryReportDTO;
import com.visoft.exceptions.SummaryReportException;
import com.visoft.services.impl.FailesSafetySummaryReportImpl;
import com.visoft.services.impl.HandlingTrafficAccidentSummaryReportImpl;
import com.visoft.services.impl.SafetySupervisoryTreatmentAfterMonitoringSummaryReportImpl;
import com.visoft.utils.SummaryReportValidationService;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;
import javax.annotation.Resource;
import java.io.IOException;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.*;

@Service(value = "summaryReportService")
public class SummaryReportService {

    @Resource private SummaryReportValidationService summaryReportValidationService;

    public StreamingResponseBody generateSummaryReport(SummaryReportDTO summaryReport) throws IOException {
        XLSXSummaryReportBuilder summaryReportBuilder;
        switch (summaryReport.getName().toUpperCase()){
            case HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT:
                summaryReportBuilder = new HandlingTrafficAccidentSummaryReportImpl();
                return summaryReportBuilder.buildXLSX(summaryReport);
            case FAILES_SAFETY_SUMMARY_REPORT:
                summaryReportBuilder = new FailesSafetySummaryReportImpl();
                return summaryReportBuilder.buildXLSX(summaryReport);
            case SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT:
                summaryReportBuilder = new SafetySupervisoryTreatmentAfterMonitoringSummaryReportImpl();
                return summaryReportBuilder.buildXLSX(summaryReport);
            default:
                throw new SummaryReportException(ERROR, summaryReportValidationService.getUnSupportedMessage(summaryReport.getName()));
        }
    }
}