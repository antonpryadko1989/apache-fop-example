package com.visoft.services.impl;

import com.visoft.dto.DeficienciesReportDTO;
import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXReportBuilder;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;

import static com.visoft.services.POIXls.addAlignmentRow;
import static com.visoft.services.POIXls.addHeaderInfo;
import static com.visoft.services.XSSFCellStyleService.*;
import static com.visoft.services.impl.ReportBodyConverter.convertReportBodyToReportBodyRows;
import static com.visoft.utils.excelReportUtils.XLSXReportUtils.*;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.*;

public class TreatmentOfImpairedByMonitoringReportImpl implements XLSXReportBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    public StreamingResponseBody buildXLSX(DeficienciesReportDTO reportDTO) throws IOException {
        boolean notEmpty = reportDTO.reportBodyListNotEmpty(reportDTO);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT_SHEET_NAME));
        setReportColumnWidths(xlsxObject.getSheet());

        final XSSFCellStyle styles[] = getTreatmentOfImpairedByMonitoringReportStyles(xlsxObject);
        final XSSFCellStyle additionalDocumentsStyles[] = getTreatmentOfImpairedByMonitoringReportStylesAdditionalDocumentsStyles(xlsxObject);

        addReportLogo(xlsxObject, reportDTO.getReportBody(), styles[0]);
        addHeaderInfo(
                xlsxObject, TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT_NAME , VERSION, DATE,
                reportDTO.getReportBody().getBodyElements().get(REPORT_NAME_VAL),
                reportDTO.getReportBody().getBodyElements().get(REPORT_VERSION_VAL),
                reportDTO.getReportBody().getBodyElements().get(REPORT_DATE_VAL),
                styles[3], styles[4], styles[5], styles[6], styles[7], styles[8]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(
                xlsxObject, reportDTO,
                styles[3], styles[14], styles[4],  styles[15], styles[9], styles[1],
                styles[2], styles[16], styles[10], styles[17], styles[7], styles[18]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addReportNumber(
                xlsxObject, FAILES_SAFETY_NUMBER,
                reportDTO.getReportBody().getBodyElements().get(REPORT_NUMBER_VAL),
                styles[11], styles[19]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addReportCreatedInfo(
                xlsxObject, reportDTO,
                styles[11], styles[13], styles[12], styles[20], styles[19]
        );
        addAlignmentRow(
                xlsxObject, SAFETY_EVENT_INFORMATION, styles[0]
        );
        addReportElementLocationSide(
                xlsxObject, reportDTO,
                styles[21], styles[4], styles[5], styles[1], styles[16]
        );
        addReportBodyRows(
                xlsxObject,
                convertReportBodyToReportBodyRows(
                        reportDTO.getReportBody().getBodyElements(),
                        TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT_BODY_KEYS,
                        TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT_BODY_VALUES),
                styles[9], styles[16], styles[10], styles[18]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addFailesSafetyReportAcceptanceOfCorrectiveAction(
                xlsxObject, reportDTO.getReportBody().getBodyElements(),
                styles[0],  styles[11], styles[13], styles[12], styles[3],
                styles[14], styles[15], styles[10], styles[17], styles[18]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addTreatmentOfImpairedByMonitoringReportAdditionalDocuments(
                xlsxObject, reportDTO.getReportBody(), notEmpty, additionalDocumentsStyles
        );
        return input2OutputService.writeWorkBook(workbook);
    }
}