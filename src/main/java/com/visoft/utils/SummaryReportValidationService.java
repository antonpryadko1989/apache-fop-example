package com.visoft.utils;

import com.visoft.dto.SummaryReportDTO;
import com.visoft.exceptions.SummaryReportException;
import org.springframework.stereotype.Service;

import java.util.HashMap;
import java.util.Map;

import static com.visoft.utils.excelReportUtils.XLSXServiceConst.NOT_ENOUGH_INFORMATION;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.NOT_SUPPORTED;

@Service(value = "summaryReportValidationService")
public class SummaryReportValidationService {

    public SummaryReportDTO validateSummaryReport(SummaryReportDTO summaryReport) {
        Map<String, String> templateValid = new HashMap<>();
        if (summaryReport.getName() == null || summaryReport.getName().equals("")) {
            templateValid.put("name", "NOT_EXIST");
        }
        if (summaryReport.getHeaders() == null) {
            templateValid.put("headers", "NOT_EXIST");
        }
        if (templateValid.isEmpty()){
            return summaryReport;
        }else throw new SummaryReportException(NOT_ENOUGH_INFORMATION, templateValid);
    }

    public Map<String, String> getUnSupportedMessage(String name){
        Map<String, String> result = new HashMap<>();
        result.put(name, NOT_SUPPORTED);
        return result;
    }
}
