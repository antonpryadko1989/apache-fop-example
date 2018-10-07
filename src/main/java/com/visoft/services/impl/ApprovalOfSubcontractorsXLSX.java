package com.visoft.services.impl;

import com.visoft.cellStyleUtil.FontParams;
import com.visoft.services.Alignment;
import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import org.apache.commons.collections4.map.LinkedMap;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;
import java.util.Map;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.CellStyleUtil.getTBMediumLRThinFont;
import static com.visoft.cellStyleUtil.CellStyleUtil.getTRBMediumLThinFont;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.services.Const.*;
import static com.visoft.services.POIXls.*;

public class ApprovalOfSubcontractorsXLSX implements XLSXBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(TemplateDTO template) throws IOException {
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());

        FontParams fontCalibri20bold = setFontParams(FONT_NAME_CALIBRI, true, (short) (20 * FONT_RATE));
        FontParams fontCalibri12bold = setFontParams(FONT_NAME_CALIBRI, true, (short) (12 * FONT_RATE));
        FontParams fontCalibri12     = setFontParams(FONT_NAME_CALIBRI, false, (short) (12 * FONT_RATE));
        XSSFCellStyle borderMediumCalibri20BoldAligmentRight = setAlignmentByParam(xlsxObject, fontCalibri20bold, Alignment.RIGHT, BorderStyle.MEDIUM);
        XSSFCellStyle lTMediumRBThinCalibri12Bold = getLTMediumRBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tMediumLRBThinCalibri12Bold = getTMediumLRBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tRMediumLBThinCalibri12Bold = getTRMediumLBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle bMediumLTRThinCalibri12Bold = getBMediumLTRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle rBMediumLTThinCalibri12Bold = getRBMediumLTThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lBMediumTRThinCalibri12Bold = getLBMediumTRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lTBMediumRThinCalibri12Bold = getLTBMediumRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tBMediumLRThinCalibri12Bold = getTBMediumLRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tRBMediumLThinCalibri12bold = getTRBMediumLThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle allBordersThinCalibri12Bold = setAllBordersThin(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lMediumTRBThinCalibri12Bold = getLMediumTRBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lBMediumRTThinCalibri12Bold = getLBMediumRTThinFont(xlsxObject, fontCalibri12bold);

        XSSFCellStyle tMediumLRBThinCalibri12 = getTMediumLRBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tRMediumLBThinCalibri12 = getTRMediumLBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle allBordersThinCalibri12 = setAllBordersThin(xlsxObject, fontCalibri12);
        XSSFCellStyle rMediumLTBThinCalibri12 = getRMediumLTBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle bMediumLTRThinCalibri12 = getBMediumLTRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle rBMediumLTThinCalibri12 = getRBMediumLTThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tBMediumLRThinCalibri12 = getTBMediumLRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tRBMediumLThinCalibri12 = getTRBMediumLThinFont(xlsxObject, fontCalibri12);

        addLogo(xlsxObject, template.getBody());
        addHeaderInfo(xlsxObject, APP_OF_SUBCONTR_TEMPLATE_NAME, VERSION, DATE,
                template.getBody().getBodyElements().get(TEMPLATE_NAME_VAL),
                template.getBody().getBodyElements().get(VERSION_VAL),
                template.getBody().getBodyElements().get(DATE_VAL),
                lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12Bold, tRMediumLBThinCalibri12Bold,
                lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12Bold, rBMediumLTThinCalibri12Bold);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(xlsxObject, template,
                lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tMediumLRBThinCalibri12Bold, tRMediumLBThinCalibri12,
                lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, allBordersThinCalibri12Bold, rMediumLTBThinCalibri12,
                lBMediumTRThinCalibri12Bold, bMediumLTRThinCalibri12, bMediumLTRThinCalibri12Bold, rBMediumLTThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addApprovalOfSubcontractorsBody(xlsxObject, template.getBody().getBodyElements(),
                lTMediumRBThinCalibri12Bold, tRMediumLBThinCalibri12, lMediumTRBThinCalibri12Bold,
                rMediumLTBThinCalibri12, lBMediumRTThinCalibri12Bold, rBMediumLTThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, CERTIFICATIONS,
                borderMediumCalibri20BoldAligmentRight);
        addCertifications(xlsxObject,
                CERTIFICATIONS_ITEM, CERTIFICATIONS_EXISTS, CERTIFICATIONS_CERTIFICATE_NO,
                CERTIFICATIONS_EXPIRATION, CERTIFICATIONS_ATTACHED_DOCUMENTS,
                lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold);
        if(notEmpty){
            addCertificationsList(xlsxObject, template.getBody().getBodyLists().get(CERTIFICATIONS_VAL),
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                    lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
        } else {
            addCertificationsList(xlsxObject, null,
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                    lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, APPROVALS, borderMediumCalibri20BoldAligmentRight);
        addApprovals(xlsxObject, APPROVALS_ROLE, APPROVALS_NAME, APPROVALS_SIGNATURE, APPROVALS_DATE, APPROVALS_STATUS,
                lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold);
        if(notEmpty){
            addApprovalsList(xlsxObject, template.getBody().getBodyLists().get(APPROVALS_VAL),
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12);
        }else {
            addApprovalsList(xlsxObject, null,
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12);
        }
        return input2OutputService.writeWorkBook(workbook);
    }

    private void addApprovalOfSubcontractorsBody(XLSXObject xlsxObject,
                                                 Map<String, String> currentMap,
                                                 XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                                                 XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6){
        setValuesToRow(xlsxObject, getApprovalOfSubcontractorsBodyRows(currentMap), style1, style2, style3, style4, style5, style6);
    }

    private static LinkedMap<String,String> getApprovalOfSubcontractorsBodyRows(Map<String, String> body) {
        LinkedMap<String, String> result = new LinkedMap<>();
        result.put(APP_OF_SUBCONTR_APPROVAL_NO, body.get(APP_OF_SUBCONTR_APPROVAL_NO_VAL));
        result.put(APP_OF_SUBCONTR_SUB_CONTRACTOR_NAME, body.get(APP_OF_SUBCONTR_SUB_CONTRACTOR_NAME_VAL));
        result.put(APP_OF_SUBCONTR_FIELD_OF_ACTIVITY, body.get(APP_OF_SUBCONTR_FIELD_OF_ACTIVITY_VAL));
        result.put(APP_OF_SUBCONTR_SUBPROJECT, body.get(APP_OF_SUBCONTR_SUBPROJECT_VAL));
        result.put(APP_OF_SUBCONTR_CONTACT, body.get(APP_OF_SUBCONTR_CONTACT_VAL));
        return result;
    }
}

