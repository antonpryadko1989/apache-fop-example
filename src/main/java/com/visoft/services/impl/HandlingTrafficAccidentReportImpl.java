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

public class HandlingTrafficAccidentReportImpl implements XLSXReportBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(DeficienciesReportDTO reportDTO) throws IOException {
        boolean notEmpty = reportDTO.reportBodyListNotEmpty(reportDTO);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(HANDLING_TRAFFIC_ACCIDENT_REPORT_SHEET_NAME));
        setReportColumnWidths(xlsxObject.getSheet());
        final XSSFCellStyle reportStyles[] = getHandlingTrafficAccidentReportStyles(xlsxObject);
        final XSSFCellStyle additionalDocumentsStyles[] = getHandlingTrafficAccidentReportAdditionalDocumentsStyles(xlsxObject);

        addReportLogo(xlsxObject, reportDTO.getReportBody(), reportStyles[0]);
        addHeaderInfo(
                xlsxObject, HANDLING_TRAFFIC_ACCIDENT_REPORT_NAME, VERSION, DATE,
                reportDTO.getReportBody().getBodyElements().get(REPORT_NAME_VAL),
                reportDTO.getReportBody().getBodyElements().get(REPORT_VERSION_VAL),
                reportDTO.getReportBody().getBodyElements().get(REPORT_DATE_VAL),
                reportStyles[3], reportStyles[4], reportStyles[5],
                reportStyles[6], reportStyles[7], reportStyles[8]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(
                xlsxObject, reportDTO,
                reportStyles[3], reportStyles[14], reportStyles[4], reportStyles[15],
                reportStyles[9], reportStyles[16], reportStyles[7], reportStyles[17]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addReportNumber(
                xlsxObject, SAFETY_DISADVANTAGES_NUMBER,
                reportDTO.getReportBody().getBodyElements().get(REPORT_NUMBER_VAL),
                reportStyles[10], reportStyles[11]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addReportCreatedInfo(
                xlsxObject, reportDTO,
                reportStyles[10], reportStyles[12] , reportStyles[11],
                reportStyles[18], reportStyles[19]
        );
        addAlignmentRow(
                xlsxObject, ACCIDENT_DETAILS, reportStyles[1]
        );
        addReportElementLocationSide(
                xlsxObject, reportDTO,
                reportStyles[21], reportStyles[4], reportStyles[5],
                reportStyles[2],  reportStyles[20]
        );
        addReportBodyRows(
                xlsxObject,
                convertReportBodyToReportBodyRows(
                        reportDTO.getReportBody().getBodyElements(),
                        HANDLING_TRAFFIC_ACCIDENT_REPORT_BODY_KEYS,
                        HANDLING_TRAFFIC_ACCIDENT_REPORT_BODY_VALUES
                        ),
                reportStyles[13], reportStyles[20], reportStyles[9], reportStyles[17]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addHandlingTrafficAccidentReportAdditionalDocuments(
                xlsxObject, reportDTO.getReportBody(), notEmpty, additionalDocumentsStyles
        );
        return input2OutputService.writeWorkBook(workbook);
    }
}