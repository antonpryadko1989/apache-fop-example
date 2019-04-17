package com.visoft.services;

import com.visoft.dto.DeficienciesReportDTO;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;

public interface XLSXReportBuilder{

    StreamingResponseBody buildXLSX(DeficienciesReportDTO deficienciesReportDTO) throws IOException;

}
