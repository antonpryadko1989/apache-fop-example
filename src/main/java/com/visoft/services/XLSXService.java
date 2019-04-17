package com.visoft.services;

import com.visoft.dto.DeficienciesReportDTO;
import com.visoft.exceptions.DeficienciesReportException;
import com.visoft.services.impl.*;
import com.visoft.utils.ReportValidationService;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.annotation.Resource;
import java.io.IOException;

import static com.visoft.utils.excelReportUtils.XLSXServiceConst.*;

@Service(value = "xlsxService")
public class XLSXService {

    @Resource
    private ReportValidationService reportValidationService;

    public StreamingResponseBody generateXLSX(DeficienciesReportDTO deficienciesReportDTO) throws IOException {
        XLSXReportBuilder reportBuilder;
        switch (deficienciesReportDTO.getReportName().toUpperCase()) {
            case HANDLING_TRAFFIC_ACCIDENT:
                reportBuilder = new HandlingTrafficAccidentReportImpl();
                return reportBuilder.buildXLSX(deficienciesReportDTO);
            case FAILES_SAFETY:
                reportBuilder = new FailesSafetyReportImpl();
                return reportBuilder.buildXLSX(deficienciesReportDTO);
            case TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT:
                reportBuilder = new TreatmentOfImpairedByMonitoringReportImpl();
                return reportBuilder.buildXLSX(deficienciesReportDTO);
            default:
                throw new DeficienciesReportException(ERROR, reportValidationService.getUnSupportedMessage(deficienciesReportDTO.getReportName()));
        }
    }
}