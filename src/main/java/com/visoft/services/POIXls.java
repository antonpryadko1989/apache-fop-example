package com.visoft.services;

import com.visoft.templates.entity.Body;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import com.visoft.templates.enums.ConfigType;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections4.map.LinkedMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.*;
import java.util.*;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.utils.Const.*;

@Service
public class POIXls {

    @Autowired
    private Input2OutputService resultService;

    public static void addLogo(XLSXObject xlsxObject,
                               Body body) {
        XSSFCellStyle style = setAllBordersByStyle(xlsxObject,
                setFontParams(FONT_NAME_CALIBRI, true,
                        (short) (20 * FONT_RATE)),
                BorderStyle.MEDIUM,
                VerticalAlignment.BOTTOM,
                HorizontalAlignment.CENTER);
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        for (int i = 1; i <= 8; i++) {
            row.createCell(i).setCellStyle(style);
        }
        row.getCell(6).setCellValue(body.getBodyElements().get(LOGO_VAL));
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 5);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 6, 8);
        row.setHeightInPoints(70);
        if(body.getLogo() != null){
            try {
                    FileInputStream is =
                            new FileInputStream(LOGO_REPOSITORY +
                                    body.getLogo().getPath());
                    byte [] bytes = org.apache.poi.util.IOUtils.toByteArray(is);
                    int pictureIndex = xlsxObject.getSheet().getWorkbook().addPicture(bytes, Workbook
                            .PICTURE_TYPE_PNG);
                    is.close();
                    CreationHelper helper = xlsxObject.getSheet().getWorkbook().getCreationHelper();
                    Drawing drawingPatriarch = xlsxObject.getSheet().createDrawingPatriarch();
                    ClientAnchor anchor = helper.createClientAnchor();

                    if(body.getLogo().getCountCells() < 3){
                        body.getLogo().setCountCells(3);
                    }
                    if(body.getLogo().getCountCells() > 8){
                        body.getLogo().setCountCells(8);
                    }
                    anchor.setCol1(9 - body.getLogo().getCountCells());
                    anchor.setRow1(xlsxObject.getRowNum());
                    anchor.setCol2(9);
                    anchor.setRow2(xlsxObject.getRowNum() + 1);
                    Picture pict = drawingPatriarch.createPicture(anchor, pictureIndex);
                } catch (IOException e) {
                    e.printStackTrace();
                }
        }
        xlsxObject.increment();
    }

    public static void addHeaderInfo(XLSXObject xlsxObject, Map<String, String> body,
                                     XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                                     XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6) {
        if(body.get(CHECKLIST_QPN_VAL)!=null){
            addHeaderInfo(xlsxObject,
                    CHECKLIST_FORM_NO, CHECKLIST_QPN, CHECKLIST_TEMPLATE_NAME, VERSION, DATE,
                    style1, style2, style3);
            addHeaderInfo(xlsxObject,
                    body.get(CHECKLIST_FORM_NO_VAL), body.get(CHECKLIST_QPN_VAL),
                    body.get(TEMPLATE_NAME_VAL), body.get(VERSION_VAL), body.get(DATE_VAL),
                    style4, style5, style6);
        }else{
            addHeaderInfo(xlsxObject,
                    CHECKLIST_FORM_NO, CHECKLIST_TEMPLATE_NAME, VERSION, DATE,
                    style1, style2, style3);
            addHeaderInfo(xlsxObject,
                    body.get(CHECKLIST_FORM_NO_VAL), body.get(TEMPLATE_NAME_VAL),
                    body.get(VERSION_VAL), body.get(DATE_VAL),
                    style4, style5, style6);
        }
    }

    public static void addCertifications (XLSXObject xlsxObject,
                                          String value1, String value2, String value3, String value4, String value5,
                                          XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3){
        setValuesToRow(xlsxObject, value1, value2, value3, value4, value5, style1, style2, style3);
    }

    public static void addCertificationsList(XLSXObject xlsxObject,
                                             List<Map<String,String>> body,
                                             XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                                             XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6,
                                             XSSFCellStyle style7, XSSFCellStyle style8, XSSFCellStyle style9,
                                             XSSFCellStyle style10, XSSFCellStyle style11, XSSFCellStyle style12) {
        if (CollectionUtils.isEmpty(body)) {
            addCertifications(xlsxObject, "", "", "", "", "", style1, style2, style3);
            addCertifications(xlsxObject, "", "", "", "", "", style7, style8, style9);
        } else {
            if(body.size()==1){
                addCertifications(xlsxObject,
                        body.get(0).get(CERTIFICATIONS_ITEM_VAL), body.get(0).get(CERTIFICATIONS_EXISTS_VAL),
                        body.get(0).get(CERTIFICATIONS_CERTIFICATE_NO_VAL), body.get(0).get(CERTIFICATIONS_EXPIRATION_VAL),
                        body.get(0).get(CERTIFICATIONS_ATTACHED_DOCUMENTS_VAL), style10, style11, style12);
            } else {
                for (int i=0; i<body.size(); i++) {
                    if (i==0) {
                        addCertifications(xlsxObject,
                                body.get(i).get(CERTIFICATIONS_ITEM_VAL), body.get(i).get(CERTIFICATIONS_EXISTS_VAL),
                                body.get(i).get(CERTIFICATIONS_CERTIFICATE_NO_VAL), body.get(i).get(CERTIFICATIONS_EXPIRATION_VAL),
                                body.get(i).get(CERTIFICATIONS_ATTACHED_DOCUMENTS_VAL), style1, style2, style3);
                    } else if (i== body.size()-1) {
                        addCertifications(xlsxObject,
                                body.get(i).get(CERTIFICATIONS_ITEM_VAL), body.get(i).get(CERTIFICATIONS_EXISTS_VAL),
                                body.get(i).get(CERTIFICATIONS_CERTIFICATE_NO_VAL), body.get(i).get(CERTIFICATIONS_EXPIRATION_VAL),
                                body.get(i).get(CERTIFICATIONS_ATTACHED_DOCUMENTS_VAL), style7, style8, style9);
                    } else {
                        addCertifications(xlsxObject,
                                body.get(i).get(CERTIFICATIONS_ITEM_VAL), body.get(i).get(CERTIFICATIONS_EXISTS_VAL),
                                body.get(i).get(CERTIFICATIONS_CERTIFICATE_NO_VAL), body.get(i).get(CERTIFICATIONS_EXPIRATION_VAL),
                                body.get(i).get(CERTIFICATIONS_ATTACHED_DOCUMENTS_VAL), style4, style5, style6);
                    }
                }
            }
        }
    }

    public static void addPreliminaryInspectionResults(XLSXObject xlsxObject,
                                                       String value1, String value2, String value3, String value4, String value5, String value6,
                                                       XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        setValuesToRow(xlsxObject, value1, value2, value3, value4, value5, value6,style1, style2, style3);
    }

    public static void addPreliminaryInspectionResultsList(XLSXObject xlsxObject, List<Map<String,String>> body,
                                                           XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                                                           XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6,
                                                           XSSFCellStyle style7, XSSFCellStyle style8, XSSFCellStyle style9,
                                                           XSSFCellStyle style10, XSSFCellStyle style11, XSSFCellStyle style12) {
        if (CollectionUtils.isEmpty(body)) {
            addPreliminaryInspectionResults(xlsxObject, "", "", "", "", "", "", style1, style2, style3);
            addPreliminaryInspectionResults(xlsxObject, "", "", "", "", "", "", style7, style8, style9);
        } else if(body.size()==1) {
            addPreliminaryInspectionResults(xlsxObject,
                    body.get(0).get(PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION_VAL), body.get(0).get(PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS_VAL),
                    body.get(0).get(PRELIM_INSPEC_RESULT_INSPECTION_RESULTS_VAL), body.get(0).get(PRELIM_INSPEC_RESULT_CERTIFICATE_NO_VAL),
                    body.get(0).get(PRELIM_INSPEC_RESULT_PASS_FAIL_VAL), body.get(0).get(PRELIM_INSPEC_RESULT_COMMENTS_VAL),
                    style10, style11, style12);
        } else {
            for (int i=0; i< body.size(); i++) {
                if (i==0) {
                    addPreliminaryInspectionResults(xlsxObject,
                            body.get(i).get(PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION_VAL), body.get(i).get(PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS_VAL),
                            body.get(i).get(PRELIM_INSPEC_RESULT_INSPECTION_RESULTS_VAL), body.get(i).get(PRELIM_INSPEC_RESULT_CERTIFICATE_NO_VAL),
                            body.get(i).get(PRELIM_INSPEC_RESULT_PASS_FAIL_VAL), body.get(i).get(PRELIM_INSPEC_RESULT_COMMENTS_VAL),
                            style1, style2, style3);
                } else if (i==body.size()-1) {
                    addPreliminaryInspectionResults(xlsxObject,
                            body.get(i).get(PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION_VAL), body.get(i).get(PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS_VAL),
                            body.get(i).get(PRELIM_INSPEC_RESULT_INSPECTION_RESULTS_VAL), body.get(i).get(PRELIM_INSPEC_RESULT_CERTIFICATE_NO_VAL),
                            body.get(i).get(PRELIM_INSPEC_RESULT_PASS_FAIL_VAL), body.get(i).get(PRELIM_INSPEC_RESULT_COMMENTS_VAL),
                            style7, style8, style9);
                } else {
                    addPreliminaryInspectionResults(xlsxObject,
                            body.get(i).get(PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION_VAL), body.get(i).get(PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS_VAL),
                            body.get(i).get(PRELIM_INSPEC_RESULT_INSPECTION_RESULTS_VAL), body.get(i).get(PRELIM_INSPEC_RESULT_CERTIFICATE_NO_VAL),
                            body.get(i).get(PRELIM_INSPEC_RESULT_PASS_FAIL_VAL), body.get(i).get(PRELIM_INSPEC_RESULT_COMMENTS_VAL),
                            style4, style5, style6);
                }
            }
        }
    }

    public static void addApprovals(XLSXObject xlsxObject,
                                    String value1, String value2, String value3, String value4, String value5,
                                    XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3){
        setValuesToRow(xlsxObject, value1, value2, value3, value4, value5, style1, style2, style3);
    }

    public static void addApprovalsList(XLSXObject xlsxObject,
                                        List<Map<String, String>> inputList,
                                        XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                                        XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6,
                                        XSSFCellStyle style7, XSSFCellStyle style8, XSSFCellStyle style9) {
        List<Map<String, String>> approvalsList = new ArrayList<>();
        if(CollectionUtils.isEmpty(inputList)){
            approvalsList.add(POC_APPROVALS_HEAD);
            approvalsList.add(POC_APPROVALS_TAIL);
        } else {
            List<Map<String, String>> approvalsQCList = new ArrayList<>();
            List<Map<String, String>> approvalsQAList = new ArrayList<>();
            List<Map<String, String>> approvalsOtherList = new ArrayList<>();
            for (Map<String, String> i : inputList) {
                if(APPROVALS_FIRST_IN_LIST_ROLE_VAL.equals(i.get(APPROVALS_ROLE_VAL))){
                    approvalsQCList.add(i);
                } else if(APPROVALS_LAST_IN_LIST_ROLE_VAL.equals(i.get(APPROVALS_ROLE_VAL))){
                    approvalsQAList.add(i);
                } else {
                    approvalsOtherList.add(i);
                }
            }
            if(CollectionUtils.isEmpty(approvalsQCList)){
                approvalsQCList.add(POC_APPROVALS_HEAD);
            }
            if(CollectionUtils.isEmpty(approvalsQAList)){
                approvalsQAList.add(POC_APPROVALS_TAIL);
            }
            approvalsList.addAll(approvalsQCList);
            approvalsList.addAll(approvalsOtherList);
            approvalsList.addAll(approvalsQAList);
        }
        for(int i=0; i<approvalsList.size(); i++){
            if (i==0) {
                addApprovals(xlsxObject,
                        approvalsList.get(i).get(APPROVALS_ROLE_VAL), approvalsList.get(i).get(APPROVALS_NAME_VAL),
                        approvalsList.get(i).get(APPROVALS_SIGNATURE_VAL), approvalsList.get(i).get(APPROVALS_DATE_VAL),
                        approvalsList.get(i).get(APPROVALS_STATUS_VAL), style1, style2, style3);
            } else if (i== approvalsList.size()-1) {
                addApprovals(xlsxObject,
                        approvalsList.get(i).get(APPROVALS_ROLE_VAL), approvalsList.get(i).get(APPROVALS_NAME_VAL),
                        approvalsList.get(i).get(APPROVALS_SIGNATURE_VAL), approvalsList.get(i).get(APPROVALS_DATE_VAL),
                        approvalsList.get(i).get(APPROVALS_STATUS_VAL), style7, style8, style9);
            } else {
                addApprovals(xlsxObject,
                        approvalsList.get(i).get(APPROVALS_ROLE_VAL), approvalsList.get(i).get(APPROVALS_NAME_VAL),
                        approvalsList.get(i).get(APPROVALS_SIGNATURE_VAL), approvalsList.get(i).get(APPROVALS_DATE_VAL),
                        approvalsList.get(i).get(APPROVALS_STATUS_VAL), style4, style5, style6);
            }
        }
    }

     public static void addDrawings(XLSXObject xlsxObject, String value1, String value2, String value3,
                                   XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        List<Integer> heightList = new ArrayList<>();
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(style1);
        row.createCell(2).setCellStyle(style1);
        row.createCell(3).setCellStyle(style1);
        row.createCell(4).setCellStyle(style2);
        row.createCell(5).setCellStyle(style2);
        row.createCell(6).setCellStyle(style3);
        row.createCell(7).setCellStyle(style3);
        row.createCell(8).setCellStyle(style3);
        row.getCell(1).setCellValue(value1);
        row.getCell(4).setCellValue(value2);
        row.getCell(6).setCellValue(value3);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 3);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 4, 5);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 6, 8);
        heightList.add(getDinamicHeight(value1, getDefLengthBCDFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        heightList.add(getDinamicHeight(value2, getDefLengthEFFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        heightList.add(getDinamicHeight(value3, getDefLengthGHIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void addDrawingsList(XLSXObject xlsxObject,
                                       List<Map<String, String>> body,
                                       XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                                       XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6,
                                       XSSFCellStyle style7, XSSFCellStyle style8, XSSFCellStyle style9,
                                       XSSFCellStyle style10, XSSFCellStyle style11, XSSFCellStyle style12) {
        if (CollectionUtils.isEmpty(body)){
            addDrawings(xlsxObject, "", "","", style1, style2, style3);
            addDrawings(xlsxObject, "", "","", style7, style8, style9);
        } else {
            if(body.size() == 1){
                addDrawings(xlsxObject,
                        body.get(0).get(DRAWING_NO_VAL),
                        body.get(0).get(DRAWING_VERSION_REVISION_VAL),
                        body.get(0).get(DRAWING_NAME_VAL),
                        style10, style11, style12);
            } else {
                for(int i=0; i<body.size(); i++){
                    if(i==0){
                        addDrawings(xlsxObject,
                                body.get(i).get(DRAWING_NO_VAL),
                                body.get(i).get(DRAWING_VERSION_REVISION_VAL),
                                body.get(i).get(DRAWING_NAME_VAL),
                                style1, style2, style3);
                    } else if(i==body.size()-1){
                        addDrawings(xlsxObject,
                                body.get(i).get(DRAWING_NO_VAL),
                                body.get(i).get(DRAWING_VERSION_REVISION_VAL),
                                body.get(i).get(DRAWING_NAME_VAL),
                                style7, style8, style9);
                    } else {
                        addDrawings(xlsxObject,
                                body.get(i).get(DRAWING_NO_VAL),
                                body.get(i).get(DRAWING_VERSION_REVISION_VAL),
                                body.get(i).get(DRAWING_NAME_VAL),
                                style4, style5, style6);
                    }
                }
            }
        }
    }

    public static void addEmptyStringWithGivenHeight(XLSXObject xlsxObject, int i) {
        xlsxObject.getSheet().createRow(xlsxObject.getRowNum()).setHeightInPoints(i);
        xlsxObject.increment();
    }

    public static void setValuesToRow(XLSXObject xlsxObject, ConfigType configType, String[] rowArray,
                                      XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        Row row = createRow(xlsxObject);
        if(configType.equals(ConfigType.ONE)){
            row.createCell(1).setCellStyle(style1);
            row.createCell(2).setCellStyle(style1);
            row.getCell(1).setCellValue(rowArray[0]);
            for (int i=3; i<8; i++){
                row.createCell(i).setCellStyle(style2);
            }
            row.createCell(8).setCellStyle(style3);
            for(int i = 1; i < rowArray.length; i++){
                row.getCell(i + 2).setCellValue(rowArray[i]);
            }
            mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                    xlsxObject.getRowNum(), 1, 2);
        } else {
            row.createCell(1).setCellStyle(style1);
            for (int i=2; i<8; i++){
                row.createCell(i).setCellStyle(style2);
            }
            row.createCell(8).setCellStyle(style3);
            for(int i = 0; i < rowArray.length; i++){
                row.getCell(i + 1).setCellValue(rowArray[i]);
            }
        }
        xlsxObject.increment();
    }

    public static void addAdditionalDocuments(XLSXObject xlsxObject,
                                              String value1, String value2, String value3, String value4, String value5,
                                              XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3){
        setValuesToRow(xlsxObject, value1, value2, value3, value4, value5, style1,style2,style3);
    }

    private static Row createRow(XLSXObject xlsxObject) {
        return xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
    }

    public static void addAdditionalDocumentsList(XLSXObject xlsxObject,
                                                  List<Map<String, String>> list,
                                                  XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                                                  XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6,
                                                  XSSFCellStyle style7, XSSFCellStyle style8, XSSFCellStyle style9,
                                                  XSSFCellStyle style10, XSSFCellStyle style11, XSSFCellStyle style12){
        if (CollectionUtils.isEmpty(list)){
            addAdditionalDocuments(xlsxObject, "", "", "", "", "", style1, style2, style3);
            addAdditionalDocuments(xlsxObject, "", "", "", "", "", style7, style8, style9);
        } else if(list.size()==1){
            addAdditionalDocuments(xlsxObject,
                    list.get(0).get(ADDITIONAL_DOCUMENTS_ITEM_VAL), list.get(0).get(ADDITIONAL_DOCUMENTS_EXISTS_VAL),
                    list.get(0).get(ADDITIONAL_DOCUMENTS_CERTIFICATE_NO_VAL), list.get(0).get(ADDITIONAL_DOCUMENTS_EXPIRATION_VAL),
                    list.get(0).get(ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS_VAL), style10, style11, style12);
        }else {
            for(int i=0; i<list.size(); i++){
                if(i==0){
                    addAdditionalDocuments(xlsxObject,
                            list.get(i).get(ADDITIONAL_DOCUMENTS_ITEM_VAL), list.get(i).get(ADDITIONAL_DOCUMENTS_EXISTS_VAL),
                            list.get(i).get(ADDITIONAL_DOCUMENTS_CERTIFICATE_NO_VAL), list.get(i).get(ADDITIONAL_DOCUMENTS_EXPIRATION_VAL),
                            list.get(i).get(ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS_VAL), style1, style2, style3);
                }else if(i==list.size()-1){
                    addAdditionalDocuments(xlsxObject,
                            list.get(i).get(ADDITIONAL_DOCUMENTS_ITEM_VAL), list.get(i).get(ADDITIONAL_DOCUMENTS_EXISTS_VAL),
                            list.get(i).get(ADDITIONAL_DOCUMENTS_CERTIFICATE_NO_VAL), list.get(i).get(ADDITIONAL_DOCUMENTS_EXPIRATION_VAL),
                            list.get(i).get(ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS_VAL), style7, style8, style9);
                }else{
                    addAdditionalDocuments(xlsxObject,
                            list.get(i).get(ADDITIONAL_DOCUMENTS_ITEM_VAL),list.get(i).get(ADDITIONAL_DOCUMENTS_EXISTS_VAL),
                            list.get(i).get(ADDITIONAL_DOCUMENTS_CERTIFICATE_NO_VAL), list.get(i).get(ADDITIONAL_DOCUMENTS_EXPIRATION_VAL),
                            list.get(i).get(ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS_VAL), style4, style5, style6);
                }
            }
        }
    }

    public static void addQaQc(XLSXObject xlsxObject,
                               String ncr, String ncrDate, String qcName, String qcDate, String qaName, String qaDate,
                               XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                               XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6,
                               XSSFCellStyle style7, XSSFCellStyle style8, XSSFCellStyle style9) {
        setValuesToRow(xlsxObject, ncr, POSITION, NAME, ncrDate, style1, style2, style3);
        Row row1 = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        setValuesToRow(xlsxObject, row1, QCM, qcName, qcDate, style4, style5, style6);
        Row row2 = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        setValuesToRow(xlsxObject, row2, QAM, qaName, qaDate, style7, style8, style9);
        row1.createCell(1).setCellStyle(style1);
        row1.createCell(2).setCellStyle(style1);
        row1.getCell(1).setCellValue(QA_QC);
        row2.createCell(1).setCellStyle(style1);
        row2.createCell(2).setCellStyle(style1);
        mergeCells(xlsxObject.getSheet(), row1.getRowNum(), row2.getRowNum(), 1, 2);
    }

    private static void setValuesToRow(XLSXObject xlsxObject, Row row,
                                       String value1, String value2, String value3,
                                       XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        int rowNum = xlsxObject.getRowNum();
        List<Integer> heightList = new ArrayList<>();
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        row.createCell(3).setCellStyle(style1);
        row.createCell(4).setCellStyle(style1);
        row.createCell(5).setCellStyle(style2);
        row.createCell(6).setCellStyle(style2);
        row.createCell(7).setCellStyle(style3);
        row.createCell(8).setCellStyle(style3);
        row.getCell(3).setCellValue(value1);
        row.getCell(5).setCellValue(value2);
        row.getCell(7).setCellValue(value3);
        mergeCells(xlsxObject.getSheet(), rowNum, rowNum, 3, 4);
        mergeCells(xlsxObject.getSheet(), rowNum, rowNum, 5, 6);
        mergeCells(xlsxObject.getSheet(), rowNum, rowNum, 7, 8);
        heightList.add(getDinamicHeight(value1, getDefLengthDEFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(value2, getDefLengthFGFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(value3, getDefLengthHIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void setValuesToRow(XLSXObject xlsxObject,
                                       String value1, String value2,
                                       XSSFCellStyle style1, XSSFCellStyle style2) {
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        row.createCell(1).setCellStyle(style1);
        row.createCell(2).setCellStyle(style1);
        for (int i=3; i<=8;i++){
            row.createCell(i).setCellStyle(style2);
        }
        row.getCell(1).setCellValue(value1);
        row.getCell(3).setCellValue(value2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 3, 8);
        heightList.add(getDinamicHeight(value2, getDefLengthDIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        row.setHeightInPoints((Collections.max(heightList)));
        xlsxObject.increment();
    }

    public static void addHeaderInfo(XLSXObject xlsxObject,
                                     String value1, String value2, String value3, String value4, String value5, String value6,
                                     XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3, XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6){
        addHeaderInfo(xlsxObject, value1, value2, value3, style1, style2, style3);
        addHeaderInfo(xlsxObject, value4, value5, value6, style4, style5, style6);
    }

    private static void addHeaderInfo(XLSXObject xlsxObject, String value1, String value2, String value3,
                                      XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        List<Integer> heightList = new ArrayList<>();
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        for(int i=1; i<=6;i++){
            row.createCell(i).setCellStyle(style1);
        }
        row.createCell(7).setCellStyle(style2);
        row.createCell(8).setCellStyle(style3);
        row.getCell(1).setCellValue(value1);
        row.getCell(7).setCellValue(value2);
        row.getCell(8).setCellValue(value3);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 6);
        heightList.add(getDinamicHeight(value1, getDefLengthBGFont18IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        heightList.add(getDinamicHeight(value2, cellLengthFont12IL[7], IL_DEF_ROW_HEIGHT_FONT_18));
        heightList.add(getDinamicHeight(value3, cellLengthFont12IL[8], IL_DEF_ROW_HEIGHT_FONT_18));
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    private static void addHeaderInfo (XLSXObject xlsxObject, String value1, String value2, String value3, String value4, String value5,
                                       XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(style1);
        row.createCell(2).setCellStyle(style2);
        for (int i=3; i <=6; i++){
            row.createCell(i).setCellStyle(style2);
        }
        row.createCell(7).setCellStyle(style2);
        row.createCell(8).setCellStyle(style3);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 3, 6);
        row.getCell(1).setCellValue(value1);
        row.getCell(2).setCellValue(value2);
        row.getCell(3).setCellValue(value3);
        row.getCell(7).setCellValue(value4);
        row.getCell(8).setCellValue(value5);
        row.setHeightInPoints(35);
        xlsxObject.increment();
    }

    private static void addHeaderInfo (XLSXObject xlsxObject, String value1, String value2, String value3, String value4,
                                       XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(style1);
        for (int i=2; i<=6; i++){
            row.createCell(i).setCellStyle(style2);
        }
        row.createCell(7).setCellStyle(style2);
        row.createCell(8).setCellStyle(style3);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 2, 6);
        row.getCell(1).setCellValue(value1);
        row.getCell(2).setCellValue(value2);
        row.getCell(7).setCellValue(value3);
        row.getCell(8).setCellValue(value4);
        row.setHeightInPoints(35);
        xlsxObject.increment();
    }

    public static void addMainInfo(XLSXObject xlsxObject, TemplateDTO template,
                                   XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3, XSSFCellStyle style4,
                                   XSSFCellStyle style5, XSSFCellStyle style6, XSSFCellStyle style7, XSSFCellStyle style8,
                                   XSSFCellStyle style9, XSSFCellStyle style10, XSSFCellStyle style11, XSSFCellStyle style12) {
        setValuesToRow(xlsxObject,
                MAIN_CONTRACTOR, template.getBody().getBodyElements().get(MAIN_CONTRACTOR_VAL),
                PROJECT_NAME, template.getBody().getBodyElements(). get(PROJECT_NAME_VAL),
                style1, style2, style3, style4);
        setValuesToRow(xlsxObject,
                MANAGEMENT_COMPANY, template.getBody().getBodyElements().get(MANAGEMENT_COMPANY_VAL),
                CONTRACT_NO, template.getBody().getBodyElements().get(CONTRACT_NO_VAL),
                style5, style6, style7, style8);
        setValuesToRow(xlsxObject,
                QC_COMPANY, template.getBody().getBodyElements().get(QC_COMPANY_VAL),
                QA_COMPANY, template.getBody().getBodyElements().get(QA_COMPANY_VAL),
                style9, style10, style11, style12);
    }


    public static void setValuesToRow(XLSXObject xlsxObject, String value1, String value2, String value3,
                                      XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(style1);
        row.createCell(2).setCellStyle(style1);
        row.createCell(3).setCellStyle(style2);
        row.createCell(4).setCellStyle(style2);
        row.createCell(5).setCellStyle(style2);
        row.createCell(6).setCellStyle(style2);
        row.createCell(7).setCellStyle(style3);
        row.createCell(8).setCellStyle(style3);
        row.getCell(1).setCellValue(value1);
        row.getCell(3).setCellValue(value2);
        row.getCell(7).setCellValue(value3);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 3, 6);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 7, 8);
        row.setHeightInPoints(30);
        xlsxObject.increment();
    }

    public static void setValuesToRow(XLSXObject xlsxObject, String value1, String value2, String value3, String value4,
                                    XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3, XSSFCellStyle style4) {
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
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
        heightList.add(getDinamicHeight(value1, getDefLengthBCFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        heightList.add(getDinamicHeight(value2, getDefLengthDEFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        heightList.add(getDinamicHeight(value3, getDefLengthFGFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        heightList.add(getDinamicHeight(value4, getDefLengthHIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void setNCRNumber(XLSXObject xlsxObject,
                                    String value1, String value2,
                                    XSSFCellStyle style1, XSSFCellStyle style2) {
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        for (int i=1; i<=4; i++){
            row.createCell(i).setCellStyle(style1);
            row.createCell(i+4).setCellStyle(style2);
        }
        row.getCell(1).setCellValue(value1);
        row.getCell(5).setCellValue(value2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 4);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 5, 8);
        heightList.add(getDinamicHeight(value1, getDefLengthBEFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        heightList.add(getDinamicHeight(value2, getDefLengthFIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void addAlignmentRow(XLSXObject xlsxObject, String value, XSSFCellStyle style) {
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 8);
        for (int i = 1; i <= 8; i++){
            row.createCell(i).setCellStyle(style);
        }
        row.createCell(1).setCellStyle(style);
        row.getCell(1).setCellValue(value);
        xlsxObject.increment();
    }

    private static void setValuesToRow(XLSXObject xlsxObject,
                                       String value1, String value2, String value3, String value4, String value5,
                                       XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        int rowNum = xlsxObject.getRowNum();
        Row row = xlsxObject.getSheet().createRow(rowNum);
        List<Integer> heightList = new ArrayList<>();
        row.createCell(1).setCellStyle(style1);
        row.createCell(2).setCellStyle(style1);
        row.createCell(3).setCellStyle(style2);
        row.createCell(4).setCellStyle(style2);
        row.createCell(5).setCellStyle(style2);
        row.createCell(6).setCellStyle(style2);
        row.createCell(7).setCellStyle(style3);
        row.createCell(8).setCellStyle(style3);
        row.getCell(1).setCellValue(value1);
        row.getCell(3).setCellValue(value2);
        row.getCell(5).setCellValue(value3);
        row.getCell(6).setCellValue(value4);
        row.getCell(7).setCellValue(value5);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 3, 4);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 7, 8);
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        heightList.add(getDinamicHeight(value1, getDefLengthBCFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(value2, getDefLengthDEFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(value3, cellLengthFont12IL[5], IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(value4, cellLengthFont12IL[6], IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(value5, getDefLengthHIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    private static void setValuesToRow(XLSXObject xlsxObject, String value1, String value2, String value3, String value4,
                                       XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        List<Integer> heightList = new ArrayList<>();
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        row.createCell(1).setCellStyle(style1);
        row.createCell(2).setCellStyle(style1);
        row.createCell(3).setCellStyle(style2);
        row.createCell(4).setCellStyle(style2);
        row.createCell(5).setCellStyle(style2);
        row.createCell(6).setCellStyle(style2);
        row.createCell(7).setCellStyle(style3);
        row.createCell(8).setCellStyle(style3);
        row.getCell(1).setCellValue(value1);
        row.getCell(3).setCellValue(value2);
        row.getCell(5).setCellValue(value3);
        row.getCell(7).setCellValue(value4);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 3, 4);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 5, 6);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 7, 8);
        heightList.add(getDinamicHeight(value1, getDefLengthBCFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(value2, getDefLengthDEFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(value3, getDefLengthFGFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(value4, getDefLengthHIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    private static void setValuesToRow (XLSXObject xlsxObject, String value1, String value2, String value3, String value4, String value5, String value6,
                                        XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3) {
        int rowNum = xlsxObject.getRowNum();
        Row row = xlsxObject.getSheet().createRow(rowNum);
        row.createCell(1).setCellStyle(style1);
        row.createCell(2).setCellStyle(style1);
        row.createCell(3).setCellStyle(style2);
        row.createCell(4).setCellStyle(style2);
        row.createCell(5).setCellStyle(style2);
        row.createCell(6).setCellStyle(style2);
        row.createCell(7).setCellStyle(style2);
        row.createCell(8).setCellStyle(style3);
        row.getCell(1).setCellValue(value1);
        row.getCell(3).setCellValue(value2);
        row.getCell(5).setCellValue(value3);
        row.getCell(6).setCellValue(value4);
        row.getCell(7).setCellValue(value5);
        row.getCell(8).setCellValue(value6);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 3, 4);
        row.setHeightInPoints(35);
        xlsxObject.increment();
    }

    public static void addChecklistElements(XLSXObject xlsxObject, List<Map<String, String>> checklistElements,
                                            XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3, XSSFCellStyle style4,
                                            XSSFCellStyle style5, XSSFCellStyle style6 ,XSSFCellStyle style7, XSSFCellStyle style8,
                                            XSSFCellStyle style9, XSSFCellStyle style10, XSSFCellStyle style11, XSSFCellStyle style12,
                                            XSSFCellStyle style13, XSSFCellStyle style14, XSSFCellStyle style15, XSSFCellStyle style16){
        if(CollectionUtils.isEmpty(checklistElements)){
            checklistElements.add(new HashMap<>());
        }
        if((checklistElements.size()%2)!=0){
            checklistElements.add(new HashMap<>());
        }
        if(checklistElements.size()==2){
            setValuesToRow(xlsxObject,
                    checklistElements.get(0).get(CHECKLIST_ELEMENT_KEY_VAL), checklistElements.get(0).get(CHECKLIST_ELEMENT_VALUE_VAL),
                    checklistElements.get(1).get(CHECKLIST_ELEMENT_KEY_VAL), checklistElements.get(1).get(CHECKLIST_ELEMENT_VALUE_VAL),
                    style13, style14, style15, style16);
        }else {
            for (int i = 0; i < checklistElements.size(); i = i + 2) {
                if (i==0) {
                    setValuesToRow(xlsxObject,
                            checklistElements.get(i).get(CHECKLIST_ELEMENT_KEY_VAL), checklistElements.get(i).get(CHECKLIST_ELEMENT_VALUE_VAL),
                            checklistElements.get(i+1).get(CHECKLIST_ELEMENT_KEY_VAL), checklistElements.get(i+1).get(CHECKLIST_ELEMENT_VALUE_VAL),
                            style1, style2, style3, style4);
                } else if (i==checklistElements.size() - 2) {
                    setValuesToRow(xlsxObject,
                            checklistElements.get(i).get(CHECKLIST_ELEMENT_KEY_VAL), checklistElements.get(i).get(CHECKLIST_ELEMENT_VALUE_VAL),
                            checklistElements.get(i+1).get(CHECKLIST_ELEMENT_KEY_VAL), checklistElements.get(i+1).get(CHECKLIST_ELEMENT_VALUE_VAL),
                            style9, style10, style11, style12);
                } else {
                    setValuesToRow(xlsxObject,
                            checklistElements.get(i).get(CHECKLIST_ELEMENT_KEY_VAL), checklistElements.get(i).get(CHECKLIST_ELEMENT_VALUE_VAL),
                            checklistElements.get(i + 1).get(CHECKLIST_ELEMENT_KEY_VAL), checklistElements.get(i + 1).get(CHECKLIST_ELEMENT_VALUE_VAL),
                            style5, style6, style7, style8);
                }
            }
        }
    }

    public static void addChecklistItem(XLSXObject xlsxObject, String value1, String value2, String value3, String value4, String value5, String value6,
                                        XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3, XSSFCellStyle style4, boolean isConst){
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(style1);
        row.createCell(2).setCellStyle(style1);
        row.createCell(3).setCellStyle(style2);
        row.createCell(4).setCellStyle(style2);
        row.createCell(5).setCellStyle(style3);
        row.createCell(6).setCellStyle(style3);
        row.createCell(7).setCellStyle(style3);
        row.createCell(8).setCellStyle(style4);
        row.getCell(1).setCellValue(value1);
        row.getCell(3).setCellValue(value2);
        row.getCell(5).setCellValue(value3);
        row.getCell(6).setCellValue(value4);
        row.getCell(7).setCellValue(value5);
        row.getCell(8).setCellValue(value6);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 3, 4);
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        if(!isConst){
            heightList.add(getDinamicHeight(value1, getDefLengthBCFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(value2, getDefLengthDEFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(value3, cellLengthFont12IL[5], IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(value4, cellLengthFont12IL[6], IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(value5, cellLengthFont12IL[7], IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(value6, cellLengthFont12IL[8], IL_DEF_ROW_HEIGHT_FONT_18));
        }
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void addChecklistItemsList(XLSXObject xlsxObject, List<Map<String, String>> list,
                                             XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3, XSSFCellStyle style4,
                                             XSSFCellStyle style5, XSSFCellStyle style6, XSSFCellStyle style7, XSSFCellStyle style8,
                                             XSSFCellStyle style9, XSSFCellStyle style10, XSSFCellStyle style11, XSSFCellStyle style12,
                                             XSSFCellStyle style13, XSSFCellStyle style14, XSSFCellStyle style15, XSSFCellStyle style16){
        if(list.isEmpty()){
            addChecklistItem(xlsxObject, "", "", "", "", "", "", style1, style2, style3, style4, false);
            addChecklistItem(xlsxObject, "", "", "", "", "", "", style9, style10, style11, style12, false);
        }
        if(list.size()==1){
            addChecklistItem(xlsxObject,
                    list.get(0).get(CHECKLIST_ITEMS_WORK_DEFINITION_VAL), list.get(0).get(CHECKLIST_ITEMS_RESPONSIBLE_PARTY_VAL),
                    list.get(0).get(CHECKLIST_ITEMS_NAME_VAL), list.get(0).get(CHECKLIST_ITEMS_SIGNATURE_VAL),
                    list.get(0).get(CHECKLIST_ITEMS_DATE_VAL), list.get(0).get(CHECKLIST_ITEMS_NOTES_VAL),
                    style13, style14, style15, style16, false);
        } else {
            for(int i=0; i<list.size(); i++){
                if(i==0){
                    addChecklistItem(xlsxObject,
                            list.get(i).get(CHECKLIST_ITEMS_WORK_DEFINITION_VAL), list.get(i).get(CHECKLIST_ITEMS_RESPONSIBLE_PARTY_VAL),
                            list.get(i).get(CHECKLIST_ITEMS_NAME_VAL), list.get(i).get(CHECKLIST_ITEMS_SIGNATURE_VAL),
                            list.get(i).get(CHECKLIST_ITEMS_DATE_VAL), list.get(i).get(CHECKLIST_ITEMS_NOTES_VAL),
                            style1, style2, style3, style4, false);
                } else if(i==list.size()-1){
                    addChecklistItem(xlsxObject,
                            list.get(i).get(CHECKLIST_ITEMS_WORK_DEFINITION_VAL), list.get(i).get(CHECKLIST_ITEMS_RESPONSIBLE_PARTY_VAL),
                            list.get(i).get(CHECKLIST_ITEMS_NAME_VAL), list.get(i).get(CHECKLIST_ITEMS_SIGNATURE_VAL),
                            list.get(i).get(CHECKLIST_ITEMS_DATE_VAL), list.get(i).get(CHECKLIST_ITEMS_NOTES_VAL),
                            style9, style10, style11, style12, false);
                } else {
                    addChecklistItem(xlsxObject,
                            list.get(i).get(CHECKLIST_ITEMS_WORK_DEFINITION_VAL), list.get(i).get(CHECKLIST_ITEMS_RESPONSIBLE_PARTY_VAL),
                            list.get(i).get(CHECKLIST_ITEMS_NAME_VAL), list.get(i).get(CHECKLIST_ITEMS_SIGNATURE_VAL),
                            list.get(i).get(CHECKLIST_ITEMS_DATE_VAL), list.get(i).get(CHECKLIST_ITEMS_NOTES_VAL),
                            style5, style6, style7, style8, false);
                }
            }
        }
    }

    public static void setValuesToRow(XLSXObject xlsxObject, LinkedMap<String, String> body,
                                      XSSFCellStyle style1, XSSFCellStyle style2,
                                      XSSFCellStyle style3, XSSFCellStyle style4,
                                      XSSFCellStyle style5, XSSFCellStyle style6){
        for (String s : body.keySet()){
            if(s.equals(body.firstKey())){
                setValuesToRow(xlsxObject, s, body.get(s), style1, style2);
            } else if(s.equals(body.lastKey())){
                setValuesToRow(xlsxObject, s, body.get(s), style5, style6);
            }else {
                setValuesToRow(xlsxObject, s, body.get(s), style3, style4);
            }
        }
    }

    public static void setValuesToRow(XLSXObject xlsxObject, LinkedMap<String, String> body,
                                      XSSFCellStyle style1, XSSFCellStyle style2,
                                      XSSFCellStyle style3, XSSFCellStyle style4){
        for (String s : body.keySet()){
            if(s.equals(body.lastKey())){
                setValuesToRow(xlsxObject, s, body.get(s), style3, style4);
            }else {
                setValuesToRow(xlsxObject, s, body.get(s), style1, style2);
            }
        }
    }

    private static void mergeCells(Sheet sheet, int fromFirstRow, int toLastRow, int fromFirstColumn, int toLastColumn) {
        sheet.addMergedRegion(new CellRangeAddress(fromFirstRow, toLastRow, fromFirstColumn, toLastColumn));
    }

    /**
     * set size for columns A - I (IL - Part)
     *
     * @param sheet     the sheet
     */
    public static void setILColumnWidths(Sheet sheet) {
        sheet.setFitToPage(true);
        sheet.setColumnWidth(0,  21 * COLUMN_WIDTHS_RATE);
        sheet.setColumnWidth(1, 152 * COLUMN_WIDTHS_RATE);
        sheet.setColumnWidth(2, 152 * COLUMN_WIDTHS_RATE);
        sheet.setColumnWidth(3, 136 * COLUMN_WIDTHS_RATE);
        sheet.setColumnWidth(4, 136 * COLUMN_WIDTHS_RATE);
        sheet.setColumnWidth(5, 152 * COLUMN_WIDTHS_RATE);
        sheet.setColumnWidth(6, 124 * COLUMN_WIDTHS_RATE);
        sheet.setColumnWidth(7, 144 * COLUMN_WIDTHS_RATE);
        sheet.setColumnWidth(8,  96 * COLUMN_WIDTHS_RATE);
    }

    /**
     * set size for columns A - ??? (RU - Part)
     *
     * @param sheet the sheet
     */
    private static Sheet setRUColumnWidths(Sheet sheet) {
        //TODO
        return sheet;
    }

//    private Workbook getWorkbook(String filePath)
//            throws IOException {
//        InputStream inputFile = new FileInputStream(filePath);
//        byte[] bytes = Files.readAllBytes(Paths.get(filePath));
//        final ByteArrayInputStream is = new ByteArrayInputStream(bytes);
//        final MediaType mediaType = new AutoDetectParser().getDetector().detect(is, new Metadata());
//        Workbook workbook = null;
//        if (mediaType.toString().equals(appXLSX)) {
//            workbook = new XSSFWorkbook(inputFile);
//            return  workbook;
//        } else if (mediaType.toString().equals(appXLS)) {
//            workbook = new HSSFWorkbook(inputFile);
//            return workbook;
//        } else {
//            throw new FileConvertException("The specified file is not Excel file");
//        }
//    }

    public static int getDinamicHeight(String value, int rowLength, int defRowHeight) {
        if(value != null){
            return getRowHeight(value, rowLength, defRowHeight)
//                    + getCountOfEnters(value) * defRowHeight
                    ;
        } else {
            return 0;
        }
    }

    private static int getRowHeight (String value, int rowLength, int defRowHeight){
        return (value.replaceAll("\n","").length()
                / rowLength + 1) * defRowHeight;
    }

    private static int getCountOfEnters(String value){
        return value.length() - value.replaceAll("\n","").length();
    }

    StreamingResponseBody addLogoToSheet(byte[] bytes, int countCells) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("LOGO");
        setILColumnWidths(sheet);
        XSSFCellStyle style = setAllBordersByBorderStyle(sheet,
                BorderStyle.MEDIUM);
        Row row = sheet.createRow(1);
        for (int i = 1; i <= 8; i++){
            row.createCell(i).setCellStyle(style);
        }
        mergeCells(sheet, 1, 1, 1, 8);
        row.setHeightInPoints(70);
        int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing drawingPatriarch = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(9 - countCells);
        anchor.setRow1(1);
        anchor.setCol2(9);
        anchor.setRow2(2);
        Picture pict = drawingPatriarch.createPicture(anchor, pictureIndex);
        return resultService.writeWorkBook(workbook);
    }
}