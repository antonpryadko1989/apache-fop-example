package com.visoft.services;

import com.visoft.dto.SummaryReportDTO;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;

public interface XLSXSummaryReportBuilder {

    StreamingResponseBody buildXLSX(SummaryReportDTO summaryReport) throws IOException;
}
