package com.visoft.services.impl;

import com.visoft.dto.SummaryReportDTO;
import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXSummaryReportBuilder;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;

import static com.visoft.services.XSSFCellStyleService.getFailesSafetySummaryReportStyles;
import static com.visoft.utils.excelReportUtils.SummaryReportConverter.changeDirectionRightToLeft;
import static com.visoft.utils.excelReportUtils.SummaryReportConverter.convertMapToSummaryReportRow;
import static com.visoft.utils.excelReportUtils.SummaryReportServiceConst.FAILES_SAFETY_SUMMARY_REPORT_ROW_VALUES;
import static com.visoft.utils.excelReportUtils.SummaryReportServiceConst.FAILES_SAFETY_SUMMARY_REPORT_TITLE;
import static com.visoft.utils.excelReportUtils.SummaryReportUtils.*;
import static com.visoft.utils.excelReportUtils.XLSXReportUtils.addEmptyStringWithGivenHeight;
import static com.visoft.utils.excelReportUtils.XLSXReportUtils.setSummaryReportColumnWidths;
import static com.visoft.utils.excelReportUtils.XLSXReportUtils.setSummaryReportSheetZomm;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.*;

public class FailesSafetySummaryReportImpl implements XLSXSummaryReportBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(SummaryReportDTO summaryReport) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XLSXObject xlsxObject = new XLSXObject(0, workbook.createSheet(FAILES_SAFETY_SUMMARY_REPORT_SHEET_NAME));
        setSummaryReportColumnWidths(xlsxObject.getSheet(), summaryReport.getName());
        setSummaryReportSheetZomm(xlsxObject, 50);
        XSSFCellStyle[] styles = getFailesSafetySummaryReportStyles(xlsxObject);
        addEmptyStringWithGivenHeight(xlsxObject, 24);
        addSummaryReportHeader(xlsxObject, 8, 11, styles[1], FAILES_SAFETY_SUMMARY_REPORT_HEADER_NAME, 55);
        addEmptyStringWithGivenHeight(xlsxObject, 24);
        addSummaryReportMainInfoRow(
                xlsxObject, 8, 9, 10, 11, 47,
                MAIN_CONTRACTOR, summaryReport.getHeaders().get(MAIN_CONTRACTOR_VAL),
                PROJECT_NAME, summaryReport.getHeaders().get(PROJECT_NAME_VAL),
                styles[2], styles[6], styles[3], styles[7]
        );
        addSummaryReportMainInfoRow(
                xlsxObject, 8, 9, 10, 11, 47,
                MANAGEMENT_COMPANY, summaryReport.getHeaders().get(MANAGEMENT_COMPANY_VAL),
                CONTRACT_NUMBER, summaryReport.getHeaders().get(CONTRACT_NUMBER_VAL),
                styles[4], styles[8], styles[5], styles[9]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 24);
        addSummaryReportRow(xlsxObject, FAILES_SAFETY_SUMMARY_REPORT_TITLE, styles[0], true);
        if (summaryReport.getValues()!=null&&!summaryReport.getValues().isEmpty()) {
            for (int i = 0; i < summaryReport.getValues().size();i++){
                addSummaryReportRow(
                        xlsxObject,
                        convertMapToSummaryReportRow(
                                summaryReport.getValues().get(i),
                                FAILES_SAFETY_SUMMARY_REPORT_ROW_VALUES,
                                i + 1),
                        styles[10], false);
            }
        }
        return input2OutputService.writeWorkBook(changeDirectionRightToLeft(workbook));
    }
}