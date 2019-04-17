package com.visoft.api;

import com.visoft.dto.SummaryReportDTO;
import com.visoft.services.SummaryReportService;
import com.visoft.utils.SummaryReportValidationService;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Controller
@RequestMapping(value = "/visoft/api")
public class SummaryReportController {

    @Value("${XLSX.contentType}")
    private String xlsxContentType;

    @RequestMapping(value = "/generateSummaryReport", method = RequestMethod.POST)
    public StreamingResponseBody generateSummaryReport (HttpServletResponse response,
                                                        @RequestBody SummaryReportDTO summaryReport) throws IOException {
        response.setCharacterEncoding("UTF-8");
        response.setContentType(xlsxContentType);
        return new SummaryReportService().generateSummaryReport(new SummaryReportValidationService().validateSummaryReport(summaryReport));
    }
}