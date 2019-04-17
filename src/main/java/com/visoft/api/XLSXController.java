package com.visoft.api;

import com.visoft.dto.DeficienciesReportDTO;
import com.visoft.services.XLSXService;
import com.visoft.utils.ReportValidationService;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.*;

@Controller
@RequestMapping(value = "/visoft/api")
public class XLSXController {

    @Value("${XLSX.contentType}")
    private String xlsxContentType;

    @RequestMapping(value = "/generateXLSX", method = RequestMethod.POST)
    public StreamingResponseBody generateXLSX(HttpServletResponse response,
                                              @RequestBody DeficienciesReportDTO deficienciesReportDTO) throws IOException {
        response.setCharacterEncoding("UTF-8");
        response.setContentType(xlsxContentType);
        return new XLSXService().generateXLSX(new ReportValidationService().validateDeficienciesReport(deficienciesReportDTO));
    }
}