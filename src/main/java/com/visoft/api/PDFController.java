package com.visoft.api;

import com.visoft.dto.DeficienciesReportDTO;
import com.visoft.services.PDFService;
import com.visoft.utils.ReportValidationService;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.servlet.http.HttpServletResponse;

@Controller
@RequestMapping(value = "/visoft/api")
public class PDFController {

    @Value("${pdf.contentType}")
    private String pdfContentType;

    @RequestMapping(value = "/generatePDF", method = RequestMethod.POST)
    public StreamingResponseBody generatePDF(HttpServletResponse response,
                                             @RequestBody DeficienciesReportDTO deficienciesReportDTO){
        response.setCharacterEncoding("UTF-8");
        response.setContentType(pdfContentType);
        return new PDFService().generatePDF(new ReportValidationService().validateDeficienciesReport(deficienciesReportDTO));
    }
}
