package com.visoft.services;

import com.visoft.exceptions.FileConvertException;
import com.visoft.services.impl.*;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;

import com.visoft.templates.enums.ConfigType;
import com.visoft.templates.enums.OrderInBlock;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.mime.MediaType;
import org.apache.tika.parser.AutoDetectParser;


import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.*;
import java.nio.file.*;
import java.util.*;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.services.Const.*;
import static com.visoft.services.Const.getDefLengthDEFont12IL;
import static com.visoft.templates.entity.TemplateDTO.getDefaultTemplate;
import static org.apache.poi.ss.util.RegionUtil.*;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

@Service
public class POIXls {

    @Value("${templates.repository}")
    private String templatesRepository;

    @Value("${app.XLSX}")
    private String appXLSX;

    @Value("${app.XLS}")
    private String appXLS;

    @Autowired
    private Input2OutputService resultService;

    public static void main(String[] args) {
//        System.out.println(getXLSX(NCR));
//        System.out.println(getXLSX(POC));
//        System.out.println(getXLSX(CHECKLIST));
//        System.out.println(getXLSX(APPROVAL_OF_SUBCONTRACTORS));
//        System.out.println(getXLSX(APPROVAL_OF_SUPPLIERS));
//        System.out.println(getXLSX(PRELIMINARY_MATERIALS_INSPECTION));
        for (int i = 0; i < 10; i++) {
//            getXLSX(NCR);
//            getXLSX(POC);
//            getXLSX(CHECKLIST);
//            getXLSX(APPROVAL_OF_SUBCONTRACTORS);
//            getXLSX(APPROVAL_OF_SUPPLIERS);
            getXLSX(PRELIMINARY_MATERIALS_INSPECTION);
        }

    }

    private static String getXLSX (String templateName){
        XLSXBuilder builder;
        switch (templateName){
            case NCR:
                builder = new NCRXlsx();
                return builder.buildXLSX(
                                getDefaultTemplate(templateName));
            case POC:
                builder = new POCXlsx();
                return builder.buildXLSX(
                                getDefaultTemplate(templateName));
            case APPROVAL_OF_SUBCONTRACTORS:
                builder = new ApprovalOfSubcontractorsXLSX();
                return builder.buildXLSX(
                                getDefaultTemplate(templateName));
            case APPROVAL_OF_SUPPLIERS:
                builder = new ApprovalOfSuppliersXlsx();
                return builder.buildXLSX(
                                getDefaultTemplate(templateName));
            case PRELIMINARY_MATERIALS_INSPECTION:
                builder = new PreliminaryMaterialsInspectionXlsx();
                return builder.buildXLSX(
                        getDefaultTemplate(templateName));
            case CHECKLIST:
                builder = new ChecklistXlsx();
                return builder.buildXLSX(
                        getDefaultTemplate(templateName));
            default:
                return "Nothing!";
        }
    }

    public static void addLogo(XLSXObject xlsxObject,
                               String logoPath,
                               String logoValue){
        XSSFCellStyle style = setAllBordersByStyle(xlsxObject,
                setFontParams(FONT_NAME_CALIBRI, true,
                        (short) (20 * FONT_RATE)),
                BorderStyle.MEDIUM,
                VerticalAlignment.BOTTOM,
                HorizontalAlignment.CENTER);
        if(logoPath!=null){
            //TODO create method for adding logo by path

        }else {
        //todo add one row with two cells, set LOGO into second cell
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(style);
        row.createCell(6).setCellStyle(style);
        row.getCell(6).setCellValue(logoValue);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 5);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 6, 8);
        setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                xlsxObject.getSheet());
        setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 6, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(70);
        }
        xlsxObject.increment();
    }

    public static void addHeaderInfo(XLSXObject xlsxObject,
                                     Map<String, String> body,
                                     XSSFCellStyle styleOne,
                                     XSSFCellStyle styleTwo,
                                     XSSFCellStyle styleThree) {
        if(body.get(CHECKLIST_QPN_VAL)!=null){
            addHeaderInfo(xlsxObject,
                    CHECKLIST_FORM_NO,
                    CHECKLIST_QPN,
                    CHECKLIST_TEMPLATE_NAME,
                    VERSION,
                    DATE,
                    styleOne, styleTwo, styleThree, true);
            addHeaderInfo(xlsxObject,
                    body.get(CHECKLIST_FORM_NO_VAL),
                    body.get(CHECKLIST_QPN_VAL),
                    body.get(TEMPLATE_NAME_VAL),
                    body.get(VERSION_VAL),
                    body.get(DATE_VAL),
                    styleOne, styleTwo, styleThree, false);
        }else{
            addHeaderInfo(xlsxObject,
                    CHECKLIST_FORM_NO,
                    CHECKLIST_TEMPLATE_NAME,
                    VERSION,
                    DATE,
                    styleOne, styleTwo, styleThree, true);
            addHeaderInfo(xlsxObject,
                    body.get(CHECKLIST_FORM_NO_VAL),
                    body.get(TEMPLATE_NAME_VAL),
                    body.get(VERSION_VAL),
                    body.get(DATE_VAL),
                    styleOne, styleTwo, styleThree, false);
        }
    }

    public static void addCertifications (XLSXObject xlsxObject,
                                          String first,
                                          String second,
                                          String third,
                                          String fourth,
                                          String fifth,
                                          XSSFCellStyle style){
        setValuesToRow(xlsxObject, first, second, third, fourth, fifth,
                style, style, true);
    }

    public static void addCertificationsList(XLSXObject xlsxObject,
                                             List<Map<String,String>> body,
                                             XSSFCellStyle styleOne,
                                             XSSFCellStyle styleTwo) {
        if (CollectionUtils.isEmpty(body)) {
            setValuesToRow(xlsxObject, null, null, null, null, null,
                    styleOne, styleTwo,false);
            setValuesToRow(xlsxObject, null, null, null, null, null,
                    styleOne, styleTwo,false);
        } else {
            for (Map<String, String> i : body) {
                setValuesToRow(xlsxObject,
                        i.get(CERTIFICATIONS_ITEM_VAL),
                        i.get(CERTIFICATIONS_EXISTS_VAL),
                        i.get(CERTIFICATIONS_CERTIFICATE_NO_VAL),
                        i.get(CERTIFICATIONS_EXPIRATION_VAL),
                        i.get(CERTIFICATIONS_ATTACHED_DOCUMENTS_VAL),
                        styleOne, styleTwo, false);
            }
        }
    }

    public static void
    addPrelimitaryInspectionResultsList(XLSXObject xlsxObject,
                                        List<Map<String,String>> body,
                                        XSSFCellStyle style) {
        if (CollectionUtils.isEmpty(body)) {
            setValuesToRow(xlsxObject, "", "", "", "", "", "",
                    style, false);
            setValuesToRow(xlsxObject, "", "", "", "", "", "",
                    style, false);
        } else {
            for (Map<String, String> i : body) {
                setValuesToRow(xlsxObject,
                        i.get(PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION_VAL),
                        i.get(PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS_VAL),
                        i.get(PRELIM_INSPEC_RESULT_INSPECTION_RESULTS_VAL),
                        i.get(PRELIM_INSPEC_RESULT_CERTIFICATE_NO_VAL),
                        i.get(PRELIM_INSPEC_RESULT_PASS_FAIL_VAL),
                        i.get(PRELIM_INSPEC_RESULT_COMMENTS_VAL),
                        style, false);
            }
        }
    }

    public static void addApprovalsList(XLSXObject xlsxObject,
                                        List<Map<String, String>>
                                                      inputList,
                                        XSSFCellStyle styleOne,
                                        XSSFCellStyle styleTwo) {
        List<Map<String, String>> approvalsList = new ArrayList<>();
        if(CollectionUtils.isEmpty(inputList)){
            approvalsList.add(POC_APPROVALS_HEAD);
            approvalsList.add(POC_APPROVALS_TAIL);
        } else {
            List<Map<String, String>> approvalsQCList = new ArrayList<>();
            List<Map<String, String>> approvalsQAList = new ArrayList<>();
            List<Map<String, String>> approvalsOtherList = new ArrayList<>();
            for (Map<String, String> i : inputList) {
                if(APPROVALS_FIRST_IN_LIST_ROLE_VAL
                        .equals(i.get(APPROVALS_ROLE_VAL))){
                    approvalsQCList.add(i);
                }else if(APPROVALS_LAST_IN_LIST_ROLE_VAL
                        .equals(i.get(APPROVALS_ROLE_VAL))){
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
        for (Map<String, String> i : approvalsList) {
            setValuesToRow(xlsxObject,
                    i.get(APPROVALS_ROLE_VAL),
                    i.get(APPROVALS_NAME_VAL),
                    i.get(APPROVALS_SIGNATURE_VAL),
                    i.get(APPROVALS_DATE_VAL),
                    i.get(APPROVALS_STATUS_VAL),
                    styleOne,
                    styleTwo,
                    false);
        }
    }

     public static void addDrawings(XLSXObject xlsxObject,
                                   String first,
                                   String second,
                                   String third,
                                   XSSFCellStyle style,
                                   boolean isConst) {
        List<Integer> heightList = new ArrayList<>();
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(style);
        row.getCell(1).setCellValue(first);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 3);
        heightList.add(getDinamicHeight(first, getDefLengthBCDFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(4).setCellStyle(style);
        row.getCell(4).setCellValue(second);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 4, 5);
        heightList.add(getDinamicHeight(second, getDefLengthEFFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(6).setCellStyle(style);
        row.getCell(6).setCellValue(third);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 6, 8);
        heightList.add(getDinamicHeight(third, getDefLengthGHIFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        if(isConst){
            setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        } else{
            setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 1, 6),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 7, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void addDrawingsList(XLSXObject xlsxObject,
                                       List<Map<String, String>> body,
                                       XSSFCellStyle style) {
        if (CollectionUtils.isEmpty(body)){
            for(int i = 0; i < 2; i++){
                addDrawings(xlsxObject, null, null, null, style, false);
            }
        }else {
            for(Map<String, String> d: body){
                addDrawings(xlsxObject,
                        d.get(DRAWING_NO_VAL),
                        d.get(DRAWING_VERSION_REVISION_VAL),
                        d.get(DRAWING_NAME_VAL),
                        style,
                        false);
            }
        }
        setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject.
                        getRowNum() - 1,xlsxObject.getRowNum() -1 , 1, 8),
                xlsxObject.getSheet());
    }

    public static String writeWorkBook(XSSFWorkbook workbook){
        try {
            FileOutputStream fileOut = new FileOutputStream(TEMPLATES +
                    "XLSX/" + System.currentTimeMillis() +".xlsx");
            workbook.write(fileOut);
            fileOut.close();
            return "OK";
        } catch (IOException e) {
            e.printStackTrace();
            return "Fale!";
        }
    }

    public static void addEmptyStringWithGivenHeight(XLSXObject xlsxObject,
                                                     int i) {
        xlsxObject.getSheet().createRow(xlsxObject.getRowNum())
                .setHeightInPoints(i);
        xlsxObject.increment();
    }

    public static void setValuesToRow(XLSXObject xlsxObject,
                                      ConfigType configType,
                                      String[] rowArray,
                                      XSSFCellStyle style,
                                      boolean isConstValue) {
        Row row = createRow(xlsxObject);
        if(configType.equals(ConfigType.ONE)){
            row.createCell(1).setCellStyle(style);
            row.getCell(1).setCellValue(rowArray[0]);
            mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                    xlsxObject.getRowNum(), 1, 2);
            for(int i = 1; i < rowArray.length; i++){
                row.createCell(i + 2).setCellStyle(style);
                row.getCell(i + 2).setCellValue(rowArray[i]);
            }
        } else {
            for(int i = 0; i < rowArray.length; i++){
                row.createCell(i + 1).setCellStyle(style);
                row.getCell(i + 1).setCellValue(rowArray[i]);
            }
        }
        if(isConstValue){
            setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                    .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                    .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        } else {
            setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                    .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 1, 2),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 3, 8),
                xlsxObject.getSheet());
        xlsxObject.increment();
    }

    public static void addAdditionalDocuments(XLSXObject xlsxObject,
                                              String first,
                                              String second,
                                              String third,
                                              String fourth,
                                              String fifth,
                                              XSSFCellStyle style){
        setValuesToRow(xlsxObject, first, second, third, fourth, fifth,
                style, style, true);
    }

    public static void addApprovals(XLSXObject xlsxObject,
                                    String first,
                                    String second,
                                    String third,
                                    String fourth,
                                    String fifth,
                                    XSSFCellStyle style){
        setValuesToRow(xlsxObject, first, second, third, fourth, fifth,
                style, style, true);
    }


    private static Row createRow(XLSXObject xlsxObject) {
        return xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
    }

    public static void addAdditionalDocumentsList(XLSXObject xlsxObject,
                                                  List<Map<String, String>> list,
                                                  XSSFCellStyle styleOne,
                                                  XSSFCellStyle styleTwo){
        if (CollectionUtils.isEmpty(list)){
            for(int i = 0; i < 2; i++){
                setValuesToRow(xlsxObject, null, null, null, null, null, styleOne,
                        styleTwo, false);
            }
        }else {
            for(Map<String, String> d: list){
                setValuesToRow(xlsxObject,
                        d.get(ADDITIONAL_DOCUMENTS_ITEM_VAL),
                        d.get(ADDITIONAL_DOCUMENTS_EXISTS_VAL),
                        d.get(ADDITIONAL_DOCUMENTS_CERTIFICATE_NO_VAL),
                        d.get(ADDITIONAL_DOCUMENTS_EXPIRATION_VAL),
                        d.get(ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS_VAL),
                        styleOne,
                        styleTwo,
                        false);
            }
        }
        setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject.
                getRowNum(),xlsxObject.getRowNum(), 1, 8),
                xlsxObject.getSheet());
    }

    public static void addQaQc(XLSXObject xlsxObject,
                               String ncr,
                               String ncrDate,
                               String qcName,
                               String qcDate,
                               String qaName,
                               String qaDate,
                               XSSFCellStyle styleOne,
                               XSSFCellStyle styleTwo) {
        Row row;
        Cell cell;
        setValuesToRow(xlsxObject, ncr, POSITION, NAME, ncrDate, styleOne);
        row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        cell = row.createCell(1);
        cell.setCellStyle(styleOne);
        cell.setCellValue(QA_QC);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum() + 1, 1, 2);
        row.createCell(0).setCellStyle(setRightBorderMedium(xlsxObject));      // don't delete
        setValuesToRow(xlsxObject, row, QCM, qcName, qcDate,
                styleOne, styleTwo);
        row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        setValuesToRow(xlsxObject, row, QAM, qaName, qaDate,
                styleOne, styleTwo);
        row.createCell(0).setCellStyle(setRightBorderMedium(xlsxObject));      // don't delete
        setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum() - 1, xlsxObject.getRowNum() - 1, 1, 8),
                xlsxObject.getSheet());
    }

    private static void setValuesToRow(XLSXObject xlsxObject,
                                       Row row,
                                       String first,
                                       String second,
                                       String third,
                                       XSSFCellStyle styleOne,
                                       XSSFCellStyle styleTwo) {
        int rowNum = xlsxObject.getRowNum();
        List<Integer> heightList = new ArrayList<>();
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        row.createCell(3).setCellStyle(styleOne);
        row.getCell(3).setCellValue(first);
        mergeCells(xlsxObject.getSheet(), rowNum, rowNum, 3, 4);
        heightList.add(getDinamicHeight(first, getDefLengthDEFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        row.createCell(5).setCellStyle(styleTwo);
        row.getCell(5).setCellValue(second);
        mergeCells(xlsxObject.getSheet(), rowNum, rowNum, 5, 6);
        heightList.add(getDinamicHeight(second, getDefLengthFGFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        row.createCell(7).setCellStyle(styleTwo);
        row.getCell(7).setCellValue(third);
        mergeCells(xlsxObject.getSheet(), rowNum, rowNum, 7, 8);
        heightList.add(getDinamicHeight(second, getDefLengthHIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        row.setHeightInPoints(Collections.max(heightList));
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 7, 8),
                xlsxObject.getSheet());
        setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 3, 8),
                xlsxObject.getSheet());
        xlsxObject.increment();
    }

    private static void setValuesToRow(XLSXObject xlsxObject,
                                       String first,
                                       String second,
                                       XSSFCellStyle borderThinCalibri12Bold,
                                       XSSFCellStyle borderThinCalibri12) {
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_12);
        row.createCell(1).setCellStyle(borderThinCalibri12Bold);
        row.getCell(1).setCellValue(first);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 2);
        heightList.add(getDinamicHeight(first, getDefLengthBCFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        row.createCell(3).setCellStyle(borderThinCalibri12);
        row.getCell(3).setCellValue(second);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 3, 8);
        heightList.add(getDinamicHeight(second, getDefLengthDIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        setBorderTop(BorderStyle.THIN, new CellRangeAddress(xlsxObject.getRowNum(),xlsxObject.getRowNum(), 1, 8), xlsxObject.getSheet());
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject.getRowNum(),xlsxObject.getRowNum(), 1, 2), xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject.getRowNum(),xlsxObject.getRowNum(), 3, 8), xlsxObject.getSheet());
        row.setHeightInPoints((Collections.max(heightList)));
        xlsxObject.increment();
    }

    public static void addHeaderInfo(XLSXObject xlsxObject,
                                     String templateName,
                                     String templateVersion,
                                     String templateDate,
                                     String templateNameVal,
                                     String templateVersionVal,
                                     String templateDateVal,
                                     XSSFCellStyle styleOne,
                                     XSSFCellStyle styleTwo){
        addHeaderInfo(xlsxObject, templateName, templateVersion, templateDate,
                true, styleOne, styleTwo);
        addHeaderInfo(xlsxObject, templateNameVal, templateVersionVal,
                templateDateVal, false, styleOne, styleTwo);
    }

    private static void addHeaderInfo(XLSXObject xlsxObject,
                                      String first,
                                      String second,
                                      String third,
                                      boolean isConstValue,
                                      XSSFCellStyle styleOne,
                                      XSSFCellStyle styleTwo) {
        List<Integer> heightList = new ArrayList<>();
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(styleOne);
        row.getCell(1).setCellValue(first);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 6);
        heightList.add(getDinamicHeight(first, getDefLengthBGFont18IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(7).setCellStyle(styleTwo);
        row.getCell(7).setCellValue(second);
        heightList.add(getDinamicHeight(second, cellLengthFont12IL[7],
                IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(8).setCellStyle(styleTwo);
        row.getCell(8).setCellValue(third);
        heightList.add(getDinamicHeight(third, cellLengthFont12IL[8],
                IL_DEF_ROW_HEIGHT_FONT_18));
        if(isConstValue){
            setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }else{
            setBorderTop(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 1, 6),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 7, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    private static void addHeaderInfo (XLSXObject xlsxObject,
                                      String first,
                                      String second,
                                      String third,
                                      String fourth,
                                      String fifth,
                                      XSSFCellStyle styleOne,
                                      XSSFCellStyle styleTwo,
                                      XSSFCellStyle styleThree,
                                      boolean isConst) {
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(styleOne);
        row.getCell(1).setCellValue(first);
        row.createCell(2).setCellStyle(styleTwo);
        row.getCell(2).setCellValue(second);
        row.createCell(3).setCellStyle(styleTwo);
        row.getCell(3).setCellValue(third);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 3, 6);
        row.createCell(7).setCellStyle(styleTwo);
        row.getCell(7).setCellValue(fourth);
        row.createCell(8).setCellStyle(styleThree);
        row.getCell(8).setCellValue(fifth);
        if(isConst){
            setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }else{
            setBorderTop(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }
        row.setHeightInPoints(35);
        xlsxObject.increment();
    }

    private static void addHeaderInfo (XLSXObject xlsxObject,
                                       String first,
                                       String second,
                                       String third,
                                       String fourth,
                                       XSSFCellStyle styleOne,
                                       XSSFCellStyle styleTwo,
                                       XSSFCellStyle styleThree,
                                       boolean isConst) {
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(styleOne);
        row.getCell(1).setCellValue(first);
        row.createCell(2).setCellStyle(styleTwo);
        row.getCell(2).setCellValue(second);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 2, 6);
        row.createCell(7).setCellStyle(styleTwo);
        row.getCell(7).setCellValue(third);
        row.createCell(8).setCellStyle(styleThree);
        row.getCell(8).setCellValue(fourth);
        if(isConst){
            setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }else{
            setBorderTop(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }
        row.setHeightInPoints(35);
        xlsxObject.increment();
    }

    public static void addMainInfo(XLSXObject xlsxObject,
                                   TemplateDTO template,
                                   XSSFCellStyle styleOne,
                                   XSSFCellStyle styleTwo) {
        addMainInfo(xlsxObject, MAIN_CONTRACTOR, template.getBody()
                        .getBodyElements().get(MAIN_CONTRACTOR_VAL),
                PROJECT_NAME, template.getBody().getBodyElements().
                        get(PROJECT_NAME_VAL),
                styleOne, styleTwo, OrderInBlock.FIRST);
        addMainInfo(xlsxObject, MANAGEMENT_COMPANY, template.getBody()
                        .getBodyElements().get(MANAGEMENT_COMPANY_VAL),
                CONTRACT_NO, template.getBody().getBodyElements()
                        .get(CONTRACT_NO_VAL), styleOne,
                styleTwo, OrderInBlock.MEDIUM);
        addMainInfo(xlsxObject, QC_COMPANY, template.getBody().getBodyElements()
                        .get(QC_COMPANY_VAL), QA_COMPANY, template.getBody()
                        .getBodyElements().get(QA_COMPANY_VAL),
                styleOne, styleTwo, OrderInBlock.LAST);
    }

    private static void addMainInfo(XLSXObject xlsxObject,
                                    String first,
                                    String second,
                                    String third,
                                    String fourth,
                                    XSSFCellStyle styleOne,
                                    XSSFCellStyle  styleTwo,
                                    OrderInBlock order) {
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        row.createCell(1).setCellStyle(styleOne);
        row.getCell(1).setCellValue(first);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 2);
        heightList.add(getDinamicHeight(first, getDefLengthBCFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(3).setCellStyle(styleTwo);
        row.getCell(3).setCellValue(second);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 3, 4);
        heightList.add(getDinamicHeight(second, getDefLengthDEFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(5).setCellStyle(styleOne);
        row.getCell(5).setCellValue(third);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 5, 6);
        heightList.add(getDinamicHeight(third, getDefLengthFGFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(7).setCellStyle(styleTwo);
        row.getCell(7).setCellValue(fourth);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 7, 8);
        heightList.add(getDinamicHeight(fourth, getDefLengthHIFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        if(order.equals(OrderInBlock.FIRST)){
            setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                    .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                    .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }
        if(order.equals(OrderInBlock.MEDIUM)){
            setBorderTop(BorderStyle.THIN, new CellRangeAddress(xlsxObject.
                    getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                    .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        }
        if(order.equals(OrderInBlock.LAST)){
            setBorderTop(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                    .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                    .getRowNum(),xlsxObject.getRowNum(), 1,8),
                    xlsxObject.getSheet());
        }
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 1, 2),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 7, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void setNCRNumber(XLSXObject xlsxObject,
                                    String first,
                                    String second,
                                    XSSFCellStyle styleOne,
                                    XSSFCellStyle styleTwo) {
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        row.createCell(1).setCellStyle(styleOne);
        row.getCell(1).setCellValue(first);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 4);
        heightList.add(getDinamicHeight(first, getDefLengthBEFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(5).setCellStyle(styleTwo);
        row.getCell(5).setCellValue(second);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 5, 8);
        heightList.add(getDinamicHeight(second, getDefLengthFIFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_18));
        setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                xlsxObject.getSheet());
        setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                xlsxObject.getSheet());
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 1, 4),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 5, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void addAlignmentRow(XLSXObject xlsxObject,
                                       String value,
                                       XSSFCellStyle style) {
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 1, 8);
        row.createCell(1).setCellStyle(style);
        row.getCell(1).setCellValue(value);
        setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject.getRowNum(),xlsxObject.getRowNum(), 1, 8), xlsxObject.getSheet());
        setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject.getRowNum(),xlsxObject.getRowNum(), 1, 8), xlsxObject.getSheet());
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject.getRowNum(),xlsxObject.getRowNum(), 1, 8), xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject.getRowNum(),xlsxObject.getRowNum(), 1, 8), xlsxObject.getSheet());
        xlsxObject.increment();
    }

    private static void setValuesToRow(XLSXObject xlsxObject,
                                       String first,
                                       String second,
                                       String third,
                                       String fourth,
                                       XSSFCellStyle style) {
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        List<Integer> heightList = new ArrayList<>();
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        row.createCell(1).setCellStyle(style);
        row.getCell(1).setCellValue(first);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 2);
        heightList.add(getDinamicHeight(first, getDefLengthBCFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_12));
        row.createCell(3).setCellStyle(style);
        row.getCell(3).setCellValue(second);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 3, 4);
        heightList.add(getDinamicHeight(second, getDefLengthDEFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_12));
        row.createCell(5).setCellStyle(style);
        row.getCell(5).setCellValue(third);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 5, 6);
        heightList.add(getDinamicHeight(third, getDefLengthFGFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        row.createCell(7).setCellStyle(style);
        row.getCell(7).setCellValue(fourth);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), 7, 8);
        heightList.add(getDinamicHeight(fourth, getDefLengthHIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_12));
        setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                xlsxObject.getSheet());
        setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                xlsxObject.getSheet());
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 1, 2),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum(),xlsxObject.getRowNum(), 7, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    private static void setValuesToRow(XLSXObject xlsxObject,
                                       String first,
                                       String second,
                                       String third,
                                       String fourth,
                                       String fifth,
                                       XSSFCellStyle styleOne,
                                       XSSFCellStyle styleTwo,
                                       boolean isConst) {
        int rowNum = xlsxObject.getRowNum();
        Row row = xlsxObject.getSheet().createRow(rowNum);
        List<Integer> heightList = new ArrayList<>();
        row.createCell(1).setCellStyle(styleOne);
        if(isConst){
            row.createCell(3).setCellStyle(styleOne);
            row.createCell(5).setCellStyle(styleOne);
            row.createCell(6).setCellStyle(styleOne);
            row.createCell(7).setCellStyle(styleOne);
            setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        } else {
            row.createCell(3).setCellStyle(styleTwo);
            row.createCell(5).setCellStyle(styleTwo);
            row.createCell(6).setCellStyle(styleTwo);
            row.createCell(7).setCellStyle(styleTwo);
            setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1,8),
                    xlsxObject.getSheet());
        }
        row.getCell(1).setCellValue(first);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 2);
        row.getCell(3).setCellValue(second);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 3, 4);
        row.getCell(5).setCellValue(third);
        row.getCell(6).setCellValue(fourth);
        row.getCell(7).setCellValue(fifth);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 7, 8);
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        heightList.add(getDinamicHeight(first, getDefLengthBCFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(second, getDefLengthDEFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(third, cellLengthFont12IL[5],
                IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(fourth, cellLengthFont12IL[6],
                IL_DEF_ROW_HEIGHT_FONT_12));
        heightList.add(getDinamicHeight(fifth, getDefLengthHIFont12IL(),
                IL_DEF_ROW_HEIGHT_FONT_12));
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 1, 2),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 7, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void setValuesToRow (XLSXObject xlsxObject,
                                       String first,
                                       String second,
                                       String third,
                                       String fourth,
                                       String fifth,
                                       String sixth,
                                       XSSFCellStyle style,
                                       boolean isConst) {
        int rowNum = xlsxObject.getRowNum();
        Row row = xlsxObject.getSheet().createRow(rowNum);
        row.createCell(1).setCellStyle(style);
        row.createCell(3).setCellStyle(style);
        row.createCell(5).setCellStyle(style);
        row.createCell(6).setCellStyle(style);
        row.createCell(7).setCellStyle(style);
        row.createCell(8).setCellStyle(style);
        if(isConst){
            setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        } else {
            setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1,8),
                    xlsxObject.getSheet());
        }
        row.getCell(1).setCellValue(first);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 2);
        row.getCell(3).setCellValue(second);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 3, 4);
        row.getCell(5).setCellValue(third);
        row.getCell(6).setCellValue(fourth);
        row.getCell(7).setCellValue(fifth);
        row.getCell(8).setCellValue(sixth);
        setBorderLeft(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 1, 2),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 7, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(35);
        xlsxObject.increment();
    }

    public static void addChecklistElements(XLSXObject xlsxObject,
                                     List<Map<String, String>>
                                             checklistElements,
                                     XSSFCellStyle styleOne,
                                     XSSFCellStyle styleTwo,
                                     XSSFCellStyle styleThree){
        if(CollectionUtils.isEmpty(checklistElements)){
            checklistElements.add(new HashMap<>());
        }
        if((checklistElements.size()%2)!=0){
            checklistElements.add(new HashMap<>());
        }
        int firstRowNum = xlsxObject.getRowNum();
        for (int i = 0; i < checklistElements.size(); i = i + 2) {
            setValuesToRow(xlsxObject,
                    checklistElements.get(i),
                    checklistElements.get(i + 1),
                    styleOne, styleTwo, styleThree);
        }
        setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(firstRowNum,
                        firstRowNum,
                        1, 8),
                xlsxObject.getSheet());
        setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum() - 1 ,xlsxObject.getRowNum() - 1, 1, 8),
                xlsxObject.getSheet());
    }

    public static void addChecklistItem(XLSXObject xlsxObject,
                                        String first,
                                        String second,
                                        String third,
                                        String fourth,
                                        String fifth,
                                        String sixth,
                                        XSSFCellStyle styleOne,
                                        XSSFCellStyle styleTwo,
                                        XSSFCellStyle styleThree,
                                        XSSFCellStyle styleFour,
                                        boolean isConst){
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(1).setCellStyle(styleOne);
        row.createCell(3).setCellStyle(styleTwo);
        row.createCell(5).setCellStyle(styleThree);
        row.createCell(6).setCellStyle(styleThree);
        row.createCell(7).setCellStyle(styleThree);
        row.createCell(8).setCellStyle(styleFour);
        row.getCell(1).setCellValue(first);
        row.getCell(3).setCellValue(second);
        row.getCell(5).setCellValue(third);
        row.getCell(6).setCellValue(fourth);
        row.getCell(7).setCellValue(fifth);
        row.getCell(8).setCellValue(sixth);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 2);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 3, 4);
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        if(isConst){
            setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
            setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                            .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                    xlsxObject.getSheet());
        } else {
            heightList.add(getDinamicHeight(first, getDefLengthBCFont12IL(),
                    IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(second, getDefLengthDEFont12IL(),
                    IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(third, cellLengthFont12IL[5],
                    IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(fourth, cellLengthFont12IL[6],
                    IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(fifth, cellLengthFont12IL[7],
                    IL_DEF_ROW_HEIGHT_FONT_18));
            heightList.add(getDinamicHeight(sixth, cellLengthFont12IL[8],
                    IL_DEF_ROW_HEIGHT_FONT_18));
        }
        setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void addChecklistItemsList(XLSXObject xlsxObject,
                                             List<Map<String, String>> list,
                                             XSSFCellStyle styleOne,
                                             XSSFCellStyle styleTwo,
                                             XSSFCellStyle styleThree,
                                             XSSFCellStyle styleFour) {
        for (Map<String, String> i: list) {
            addChecklistItem(xlsxObject,
                    i.get(CHECKLIST_ITEMS_WORK_DEFINITION_VAL),
                    i.get(CHECKLIST_ITEMS_RESPONSIBLE_PARTY_VAL),
                    i.get(CHECKLIST_ITEMS_NAME_VAL),
                    i.get(CHECKLIST_ITEMS_SIGNATURE_VAL),
                    i.get(CHECKLIST_ITEMS_DATE_VAL),
                    i.get(CHECKLIST_ITEMS_NOTES_VAL),
                    styleOne,
                    styleTwo,
                    styleThree,
                    styleFour,
                    false);
            System.out.println(i);
        }
        setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum() - 1,xlsxObject.getRowNum() - 1, 1, 8),
                xlsxObject.getSheet());
    }

    private static void setValuesToRow (XLSXObject xlsxObject,
                                        Map<String, String> firstPair,
                                        Map<String, String> secondPair,
                                        XSSFCellStyle styleOne,
                                        XSSFCellStyle styleTwo,
                                        XSSFCellStyle styleThree) {
        List<Integer> heightList = new ArrayList<>();
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        heightList.add(IL_DEF_ROW_HEIGHT_FONT_18);
        row.createCell(1).setCellStyle(styleOne);
        row.getCell(1).setCellValue(firstPair.get(CHECKLIST_ELEMENT_KEY_VAL));
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 1, 2);
        heightList.add(getDinamicHeight(firstPair.get(CHECKLIST_ELEMENT_KEY_VAL),
                getDefLengthBCFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(3).setCellStyle(styleTwo);
        row.getCell(3).setCellValue(firstPair.get(CHECKLIST_ELEMENT_VALUE_VAL));
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 3, 4);
        heightList.add(getDinamicHeight(firstPair.get(CHECKLIST_ELEMENT_VALUE_VAL),
                getDefLengthDEFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(5).setCellStyle(styleThree);
        row.getCell(5).setCellValue(secondPair.get(CHECKLIST_ELEMENT_KEY_VAL));
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 5, 6);
        heightList.add(getDinamicHeight(secondPair.get(CHECKLIST_ELEMENT_KEY_VAL),
                getDefLengthFGFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        row.createCell(7).setCellStyle(styleTwo);
        row.getCell(7).setCellValue(secondPair.get(CHECKLIST_ELEMENT_VALUE_VAL));
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(),
                xlsxObject.getRowNum(), 7, 8);
        heightList.add(getDinamicHeight(secondPair.get
                        (CHECKLIST_ELEMENT_VALUE_VAL),
                getDefLengthHIFont12IL(), IL_DEF_ROW_HEIGHT_FONT_18));
        setBorderBottom(BorderStyle.THIN, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 1, 8),
                xlsxObject.getSheet());
        setBorderRight(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                        .getRowNum(),xlsxObject.getRowNum(), 7, 8),
                xlsxObject.getSheet());
        row.setHeightInPoints(Collections.max(heightList));
        xlsxObject.increment();
    }

    public static void setValuesToRow(XLSXObject xlsxObject,
                                      Map<String, String> body,
                                      XSSFCellStyle styleOne,
                                      XSSFCellStyle styleTwo){
        int firstRowNum = xlsxObject.getRowNum();
        for (String s : body.keySet()){
            setValuesToRow(xlsxObject, s, body.get(s), styleOne, styleTwo);
        }
        setBorderTop(BorderStyle.MEDIUM, new CellRangeAddress(firstRowNum,
                        firstRowNum,
                        1, 8),
                xlsxObject.getSheet());
        setBorderBottom(BorderStyle.MEDIUM, new CellRangeAddress(xlsxObject
                .getRowNum() - 1 ,xlsxObject.getRowNum() - 1, 1, 8),
                xlsxObject.getSheet());
    }

    private static void mergeCells(Sheet sheet,
                                   int fromFirstRow,
                                   int toLastRow,
                                   int fromFirstColumn,
                                   int toLastColumn) {
        sheet.addMergedRegion(
                new CellRangeAddress(
                        fromFirstRow,
                        toLastRow,
                        fromFirstColumn,
                        toLastColumn)
        );
    }

    /**
     * set size for columns A - I (IL - Part)
     *
     * @param sheet     the sheet
     */
    public static void setILColumnWidths(Sheet sheet) {
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

    private Workbook getWorkbook(String filePath)
            throws IOException {
        InputStream inputFile = new FileInputStream(filePath);
        byte[] bytes = Files.readAllBytes(Paths.get(filePath));
        final ByteArrayInputStream is = new ByteArrayInputStream(bytes);
        final MediaType mediaType = new AutoDetectParser().getDetector().detect(is, new Metadata());
        Workbook workbook = null;
        if (mediaType.toString().equals(appXLSX)) {
            workbook = new XSSFWorkbook(inputFile);
            return  workbook;
        } else if (mediaType.toString().equals(appXLS)) {
            workbook = new HSSFWorkbook(inputFile);
            return workbook;
        } else {
            throw new FileConvertException("The specified file is not Excel file");
        }
    }

    private static int getDinamicHeight(String value, int rowLength, int defRowHeight) {
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

    StreamingResponseBody getExcelFile(TemplateDTO template) throws IOException {
//        File f = Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()).toFile();
//        if (f.isDirectory()) {
//            throw new PathValidationException("error", "It's directory");
//        }
//        if (!f.exists()) {
//            throw new PathValidationException("error", "File not exist");
//        }
//        Workbook workbook = getWorkbook(Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()).toString());
//        Sheet sheet = workbook.getSheetAt(0);
//        for (Row row : sheet) {
//            Iterator<Cell> cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
//                    String newCellValueKey = TemplateValue.isValue(cell.getStringCellValue());
//                    if (newCellValueKey.length() > 0) {
//                        cell.setCellValue(template.getTemplateBody().get(newCellValueKey));
//                    }
//                }
//            }
//        }
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        workbook.write(out);
//        InputStream inputStream = new ByteArrayInputStream(out.toByteArray());
//        return resultService.getOutput(inputStream);
        return null;
    }
}