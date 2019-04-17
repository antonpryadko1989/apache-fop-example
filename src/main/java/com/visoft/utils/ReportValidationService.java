package com.visoft.utils;

import com.visoft.dto.DeficienciesReportDTO;
import com.visoft.exceptions.DeficienciesReportException;
import org.springframework.stereotype.Service;

import java.util.HashMap;
import java.util.Map;

import static com.visoft.utils.excelReportUtils.XLSXServiceConst.NOT_ENOUGH_INFORMATION;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.NOT_SUPPORTED;

@Service(value = "reportValidationService")
public class ReportValidationService {

    public DeficienciesReportDTO validateDeficienciesReport(DeficienciesReportDTO reportDTO) {
        Map<String, String> templateValid = new HashMap<>();
        if (reportDTO.getReportName() == null || reportDTO.getReportName().equals("")) {
            templateValid.put("reportName", "NOT_EXIST");
        }
//        if (reportDTO.getOutputName() == null || reportDTO.getOutputName().equals("")){
//            templateValid.put("outputName", "NOT_EXIST");
//        }
        if (reportDTO.getReportBody() == null) {
            templateValid.put("reportBody", "NOT_EXIST");
        }
        if (templateValid.isEmpty()){
            return reportDTO;
        }else
            throw new DeficienciesReportException(NOT_ENOUGH_INFORMATION, templateValid);
    }

    public Map<String, String> getUnSupportedMessage(String name){
        Map<String, String> result = new HashMap<>();
        result.put(name, NOT_SUPPORTED);
        return result;
    }
}
