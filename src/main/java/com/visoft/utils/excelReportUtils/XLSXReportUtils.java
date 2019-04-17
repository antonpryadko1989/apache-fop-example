package com.visoft.utils.excelReportUtils;

import com.visoft.dto.DeficienciesReportDTO;
import com.visoft.dto.ReportBody;
import com.visoft.dto.ReportLogo;
import com.visoft.templates.entity.XLSXObject;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.map.LinkedMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import static com.visoft.services.POIXls.addAlignmentRow;
import static com.visoft.services.POIXls.getDinamicHeight;
import static com.visoft.services.POIXls.setValuesToRow;
import static com.visoft.utils.Const.*;
import static com.visoft.utils.excelReportUtils.ReportLogoValidator.isValidReportLogo;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.*;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.ADDITIONAL_DOCUMENTS_VAL;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.COLUMN_WIDTHS_RATE;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.ELEMENT;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.ELEMENT_VAL;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.MAIN_CONTRACTOR;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.MAIN_CONTRACTOR_VAL;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.MANAGEMENT_COMPANY;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.MANAGEMENT_COMPANY_VAL;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.PROJECT_NAME;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.PROJECT_NAME_VAL;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.QA_COMPANY;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.QA_COMPANY_VAL;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.QC_COMPANY;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.QC_COMPANY_VAL;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.SIDE;
import static com.visoft.utils.excelReportUtils.XLSXServiceConst.SIDE_VAL;

public class XLSXReportUtils {

    public static void setReportColumnWidths(Sheet sheet) {
        sheet.setFitToPage(true);
        int[] columns = {21, 152, 152, 136, 136, 152, 124, 144, 96};
        for (int i = 0; i < columns.length; i++) {
            sheet.setColumnWidth(i,  columns[i] * COLUMN_WIDTHS_RATE);
        }
    }

    public static void setSummaryReportColumnWidths(Sheet sheet, String summaryReportName) {
        sheet.setFitToPage(true);
        int[] columns;
        switch (summaryReportName){
            case HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT:
                columns = HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT_COLUMN_WIDTHS;
                break;
            case FAILES_SAFETY_SUMMARY_REPORT:
                columns = FAILES_SAFETY_SUMMARY_REPORT_COLUMN_WIDTHS;
                break;
            case SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT:
                columns = SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT_COLUMN_WIDTHS;
                break;
            default:
                columns = DEFAULT_REPORT_COLUMN_WIDTHS;
                break;
        }
        for (int i = 0; i < columns.length; i++) {
            sheet.setColumnWidth(i,  columns[i] * COLUMN_WIDTHS_RATE);
        }
    }

    public static void setSummaryReportSheetZomm(XLSXObject xlsxObject, int zoom){
        xlsxObject.getSheet().setZoom(zoom);
    }

    public static void addReportLogo(XLSXObject xlsxObject, ReportBody reportBody, XSSFCellStyle style){
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        for (int i = 1; i <= 8; i++) {
            row.createCell(i).setCellStyle(style);
        }
        row.getCell(6).setCellValue(reportBody.getBodyElements().get(MAIN_CONTRACTOR_VAL));
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 5);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 6, 8);
        row.setHeightInPoints(70);
        if (reportBody.getReportLogo() != null) {
            if(isValidReportLogo(reportBody.getReportLogo())){
                int pictureIndex = xlsxObject.getSheet().getWorkbook().addPicture(reportBody.getReportLogo().getData(), Workbook.PICTURE_TYPE_PNG);
                CreationHelper helper = xlsxObject.getSheet().getWorkbook().getCreationHelper();
                Drawing drawingPatriarch = xlsxObject.getSheet().createDrawingPatriarch();
                ClientAnchor anchor = helper.createClientAnchor();
                if (reportBody.getReportLogo().getCountCells() < 3) {
                    reportBody.getReportLogo().setCountCells(3);
                }
                if (reportBody.getReportLogo().getCountCells() > 8) {
                    reportBody.getReportLogo().setCountCells(8);
                }
                anchor.setCol1(9 - reportBody.getReportLogo().getCountCells());
                anchor.setRow1(xlsxObject.getRowNum());
                anchor.setCol2(9);
                anchor.setRow2(xlsxObject.getRowNum() + 1);
                drawingPatriarch.createPicture(anchor, pictureIndex);
            }
        }
        xlsxObject.increment();
    }

    static void mergeCells(Sheet sheet, int fromFirstRow, int toLastRow, int fromFirstColumn, int toLastColumn) {
        sheet.addMergedRegion(
                new CellRangeAddress(
                        fromFirstRow,
                        toLastRow,
                        fromFirstColumn,
                        toLastColumn
                )
        );
    }

    public static void addEmptyStringWithGivenHeight(XLSXObject xlsxObject, int i) {
        xlsxObject.getSheet().createRow(xlsxObject.getRowNum()).setHeightInPoints(i);
        xlsxObject.increment();
    }

    public static void addMainInfo(XLSXObject xlsxObject, DeficienciesReportDTO reportDTO,
                                   XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3, XSSFCellStyle style4,
                                   XSSFCellStyle style5, XSSFCellStyle style6, XSSFCellStyle style7, XSSFCellStyle style8) {
        setValuesToRow(xlsxObject,
                MAIN_CONTRACTOR, reportDTO.getReportBody().getBodyElements().get(MAIN_CONTRACTOR_VAL),
                PROJECT_NAME, reportDTO.getReportBody().getBodyElements(). get(PROJECT_NAME_VAL),
                style1, style2, style3, style4);
        setValuesToRow(xlsxObject,
                MANAGEMENT_COMPANY, reportDTO.getReportBody().getBodyElements().get(MANAGEMENT_COMPANY_VAL),
                CONTRACT_NUMBER, reportDTO.getReportBody().getBodyElements().get(CONTRACT_NUMBER_VAL),
                style5, style6, style7, style8);
    }

    public static void addMainInfo(XLSXObject xlsxObject, DeficienciesReportDTO reportDTO,
                                   XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3, XSSFCellStyle style4,
                                   XSSFCellStyle style5, XSSFCellStyle style6, XSSFCellStyle style7, XSSFCellStyle style8,
                                   XSSFCellStyle style9, XSSFCellStyle style10, XSSFCellStyle style11, XSSFCellStyle style12) {
        setValuesToRow(xlsxObject,
                MAIN_CONTRACTOR, reportDTO.getReportBody().getBodyElements().get(MAIN_CONTRACTOR_VAL),
                PROJECT_NAME, reportDTO.getReportBody().getBodyElements(). get(PROJECT_NAME_VAL),
                style1, style2, style3, style4);
        setValuesToRow(xlsxObject,
                MANAGEMENT_COMPANY, reportDTO.getReportBody().getBodyElements().get(MANAGEMENT_COMPANY_VAL),
                CONTRACT_NUMBER, reportDTO.getReportBody().getBodyElements().get(CONTRACT_NUMBER_VAL),
                style5, style6, style7, style8);
        setValuesToRow(xlsxObject,
                QC_COMPANY, reportDTO.getReportBody().getBodyElements().get(QC_COMPANY_VAL),
                QA_COMPANY, reportDTO.getReportBody().getBodyElements().get(QA_COMPANY_VAL),
                style9, style10, style11, style12);
    }

    public static void addReportNumber(XLSXObject xlsxObject, String key, String value, XSSFCellStyle style1, XSSFCellStyle style2){
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        for (int i = 1; i <= 4 ; i++) {
            row.createCell(i).setCellStyle(style1);
        }
        for (int i = 5; i <= 8 ; i++) {
            row.createCell(i).setCellStyle(style2);
        }
        row.getCell(1).setCellValue(key);
        row.getCell(5).setCellValue(value);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 4);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 5, 8);
        heightList.add(getDinamicHeight(key, getDefLengthBEFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        heightList.add(getDinamicHeight(value, getDefLengthFIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void addReportCreatedInfo(XLSXObject xlsxObject,
                                            DeficienciesReportDTO reportDTO,
                                            XSSFCellStyle style1,
                                            XSSFCellStyle style2,
                                            XSSFCellStyle style3,
                                            XSSFCellStyle style4,
                                            XSSFCellStyle style5){
        setValuesToReportCreatedInfo(xlsxObject,
                ROLE, OPENED_BY, OPENED_DATE, OPENED_HOUR,
                style1, style2, style2, style3);
        setValuesToReportCreatedInfo(xlsxObject,
                reportDTO.getReportBody().getBodyElements(). get(ROLE_VAL),
                reportDTO.getReportBody().getBodyElements(). get(OPENED_BY_VAL),
                reportDTO.getReportBody().getBodyElements(). get(OPENED_DATE_VAL),
                reportDTO.getReportBody().getBodyElements(). get(OPENED_HOUR_VAL),
                style1, style2, style4, style5);
    }

    private static void setValuesToReportCreatedInfo(XLSXObject xlsxObject,
                                                     String value1,
                                                     String value2,
                                                     String value3,
                                                     String value4,
                                                     XSSFCellStyle style1,
                                                     XSSFCellStyle style2,
                                                     XSSFCellStyle style3,
                                                     XSSFCellStyle style4){
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(style1);
        row.createCell(2).setCellStyle(style1);
        row.createCell(3).setCellStyle(style2);
        row.createCell(4).setCellStyle(style2);
        row.createCell(5).setCellStyle(style3);
        row.createCell(6).setCellStyle(style3);
        row.createCell(7).setCellStyle(style4);
        row.createCell(8).setCellStyle(style4);
        row.getCell(1).setCellValue(value1);
        row.getCell(3).setCellValue(value2);
        row.getCell(5).setCellValue(value3);
        row.getCell(7).setCellValue(value4);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 3, 4);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 5, 6);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 7, 8);
        row.setHeightInPoints(18);
        xlsxObject.increment();
    }

    public static void addReportElementLocationSide(XLSXObject xlsxObject,
                                                    DeficienciesReportDTO reportDTO,
                                                    XSSFCellStyle style1,
                                                    XSSFCellStyle style2,
                                                    XSSFCellStyle style3,
                                                    XSSFCellStyle style4,
                                                    XSSFCellStyle style5){
        Row row1 = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        setValuesToReportElementLocationSide(xlsxObject, row1, ELEMENT, LOCATION, SIDE,
                style2, style3, 18);
        Row row2 = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        setValuesToReportElementLocationSide(xlsxObject, row2,
                reportDTO.getReportBody().getBodyElements(). get(ELEMENT_VAL),
                reportDTO.getReportBody().getBodyElements(). get(LOCATION_VAL),
                reportDTO.getReportBody().getBodyElements(). get(SIDE_VAL),
                style4, style5, 30
                );
        row1.createCell(1).setCellStyle(style1);
        row1.createCell(2).setCellStyle(style1);
        row2.createCell(1).setCellStyle(style1);
        row2.createCell(2).setCellStyle(style1);
        mergeCells(xlsxObject.getSheet(), row1.getRowNum(), row2.getRowNum(), 1, 2);
    }

    private static void setValuesToReportElementLocationSide(XLSXObject xlsxObject,
                                                             Row row,
                                                             String value1,
                                                             String value2,
                                                             String value3,
                                                             XSSFCellStyle style1,
                                                             XSSFCellStyle style2,
                                                             int rowHeight){
        int rowNum = xlsxObject.getRowNum();
        row.createCell(3).setCellStyle(style1);
        row.createCell(4).setCellStyle(style1);
        row.createCell(5).setCellStyle(style1);
        row.createCell(6).setCellStyle(style1);
        row.createCell(7).setCellStyle(style2);
        row.createCell(8).setCellStyle(style2);
        row.getCell(3).setCellValue(value1);
        row.getCell(5).setCellValue(value2);
        row.getCell(7).setCellValue(value3);
        mergeCells(xlsxObject.getSheet(), rowNum, rowNum, 3, 4);
        mergeCells(xlsxObject.getSheet(), rowNum, rowNum, 5, 6);
        mergeCells(xlsxObject.getSheet(), rowNum, rowNum, 7, 8);
        row.setHeightInPoints(rowHeight);
        xlsxObject.increment();
    }

    public static void addReportBodyRows(XLSXObject xlsxObject,
                                         LinkedMap<String, String> currentMap,
                                         XSSFCellStyle style1,
                                         XSSFCellStyle style2,
                                         XSSFCellStyle style3,
                                         XSSFCellStyle style4){
        for (String s : currentMap.keySet()){
            if(s.equals(currentMap.lastKey())){
                setValuesToRow(xlsxObject, s, currentMap.get(s), style3, style4);
            }else {
                setValuesToRow(xlsxObject, s, currentMap.get(s), style1, style2);
            }
        }
    }




    public static void addFailesSafetyReportAcceptanceOfCorrectiveAction(XLSXObject xlsxObject,
                                                                         Map<String, String> reportBodyElements,
                                                                         XSSFCellStyle style1,
                                                                         XSSFCellStyle style2,
                                                                         XSSFCellStyle style3,
                                                                         XSSFCellStyle style4,
                                                                         XSSFCellStyle style5,
                                                                         XSSFCellStyle style6,
                                                                         XSSFCellStyle style7,
                                                                         XSSFCellStyle style8,
                                                                         XSSFCellStyle style9,
                                                                         XSSFCellStyle style10){
        addAlignmentRow(xlsxObject, FAILES_SAFETY_ACCEPTANCE_OF_CORRECTIVE_ACTION, style1);
        setValuesToRow(xlsxObject,
                ACCEPTANCE_OF_CORRECTIVE_ACTION_ROLE,
                ACCEPTANCE_OF_CORRECTIVE_ACTION_NAME, ACCEPTANCE_OF_CORRECTIVE_ACTION_DATE,
                style2, style3, style4);
        setValuesToRow(xlsxObject,
                ACCEPTANCE_OF_CORRECTIVE_ACTION_ROLE_QC,
                reportBodyElements.get(QC_NAME_VAL), reportBodyElements.get(QC_APPROVAL_DATE_VAL)
                , style5, style6, style7);
        setValuesToRow(xlsxObject, ACCEPTANCE_OF_CORRECTIVE_ACTION_ROLE_FOREMAN,
                reportBodyElements.get(FOREMAN_NAME_VAL), reportBodyElements.get(FOREMAN_APPROVAL_DATE_VAL)
                , style8, style9, style10);
    }

    public static void addHandlingTrafficAccidentReportAdditionalDocuments(XLSXObject xlsxObject,
                                                                           ReportBody reportBody,
                                                                           boolean notEmpty,
                                                                           XSSFCellStyle styles[]){
        addAlignmentRow(xlsxObject, HANDLING_TRAFFIC_ACCIDENT_ADDITIONAL_DOCUMENTS, styles[0]);
        setValuesToRow(xlsxObject,
                HANDLING_TRAFFIC_ACCIDENT_ADDITIONAL_DOCUMENT_TYPE,
                HANDLING_TRAFFIC_ACCIDENT_ADDITIONAL_DOCUMENT_DESCRIPTION,
                styles[1], styles[2]);
        if(notEmpty){
            addReportAdditionalDocumentsList(xlsxObject, reportBody.getBodyLists().get(ADDITIONAL_DOCUMENTS_VAL), styles);
        }else{
            setTwoEmptyRows(xlsxObject, styles[5], styles[6],styles[9], styles[10]);
        }
    }

    public static void addFailesSafetyReportAdditionalDocuments(XLSXObject xlsxObject,
                                                                ReportBody reportBody,
                                                                boolean notEmpty,
                                                                XSSFCellStyle styles[]){
        addAlignmentRow(xlsxObject, FAILES_SAFETY_ADDITIONAL_DOCUMENTS, styles[0]);
        setValuesToRow(xlsxObject,
                FAILES_SAFETY_ADDITIONAL_DOCUMENT_TYPE,
                FAILES_SAFETY_ADDITIONAL_DOCUMENT_DESCRIPTION,
                styles[1], styles[2]);
        if(notEmpty){
            addReportAdditionalDocumentsList(xlsxObject, reportBody.getBodyLists().get(ADDITIONAL_DOCUMENTS_VAL), styles);
        }else{
            setTwoEmptyRows(xlsxObject, styles[5], styles[6],styles[9], styles[10]);
        }
    }

    public static void addTreatmentOfImpairedByMonitoringReportAdditionalDocuments(XLSXObject xlsxObject,
                                                                                   ReportBody reportBody,
                                                                                   boolean notEmpty,
                                                                                   XSSFCellStyle styles[]){
        addAlignmentRow(xlsxObject, TREATMENT_OF_IMPAIRED_BY_MONITORING_ADDITIONAL_DOCUMENTS, styles[0]);
        setValuesToRow(xlsxObject,
                TREATMENT_OF_IMPAIRED_BY_MONITORING_ADDITIONAL_DOCUMENT_TYPE,
                TREATMENT_OF_IMPAIRED_BY_MONITORING_ADDITIONAL_DOCUMENT_DESCRIPTION,
                styles[1], styles[2]);
        if(notEmpty){
            addReportAdditionalDocumentsList(xlsxObject, reportBody.getBodyLists().get(ADDITIONAL_DOCUMENTS_VAL), styles);
        }else{
            setTwoEmptyRows(xlsxObject, styles[5], styles[6],styles[9], styles[10]);
        }
    }

    private static void setTwoEmptyRows(XLSXObject xlsxObject,
                                        XSSFCellStyle style1,
                                        XSSFCellStyle style2,
                                        XSSFCellStyle style3,
                                        XSSFCellStyle style4){
        setValuesToRow(xlsxObject, "", "", style1, style2);
        setValuesToRow(xlsxObject, "", "", style3, style4);
    }

    private static void addReportAdditionalDocumentsList(XLSXObject xlsxObject,
                                                         List<Map<String, String>> additionalDocuments,
                                                         XSSFCellStyle styles[]){
        if(!CollectionUtils.isEmpty(additionalDocuments)){
            if(additionalDocuments.size()==1){
                addOneReportAdditionalDocument(xlsxObject, additionalDocuments.get(0), styles[3], styles[4]);
            }else if(additionalDocuments.size()==2){
                addTwoReportAdditionalDocuments(xlsxObject, additionalDocuments.get(0), additionalDocuments.get(1), styles[5], styles[6],styles[9], styles[10]);
            }else{
                addReportAdditionalDocuments(xlsxObject, additionalDocuments, styles);
            }
        }else {
            setTwoEmptyRows(xlsxObject, styles[5], styles[6],styles[9], styles[10]);
        }
    }

    private static void addOneReportAdditionalDocument(XLSXObject xlsxObject,
                                                       Map<String, String> additionalDocument,
                                                       XSSFCellStyle style1,
                                                       XSSFCellStyle style2){
        setValuesToRow(
                xlsxObject,
                additionalDocument.get(ADDITIONAL_DOCUMENT_TYPE_VAL),
                additionalDocument.get(ADDITIONAL_DOCUMENT_DESCRIPTION_VAL),
                style1,
                style2
        );
    }

    private static void addTwoReportAdditionalDocuments(XLSXObject xlsxObject,
                                                        Map<String, String> additionalDocument1,
                                                        Map<String, String> additionalDocument2,
                                                        XSSFCellStyle style1,
                                                        XSSFCellStyle style2,
                                                        XSSFCellStyle style3,
                                                        XSSFCellStyle style4){
        addOneReportAdditionalDocument(xlsxObject, additionalDocument1, style1, style2);
        addOneReportAdditionalDocument(xlsxObject, additionalDocument2, style3, style4);
    }

    private static void addReportAdditionalDocuments(XLSXObject xlsxObject,
                                                     List<Map<String, String>> additionalDocuments,
                                                     XSSFCellStyle styles[]){
        for (int i = 0; i < additionalDocuments.size(); i++) {
            if (i==0){
                addOneReportAdditionalDocument(xlsxObject, additionalDocuments.get(i), styles[5], styles[6]);
            }else if(i==additionalDocuments.size()-1){
                addOneReportAdditionalDocument(xlsxObject, additionalDocuments.get(i), styles[9], styles[10]);
            }else {
                addOneReportAdditionalDocument(xlsxObject, additionalDocuments.get(i), styles[7], styles[8]);
            }
        }
    }
}