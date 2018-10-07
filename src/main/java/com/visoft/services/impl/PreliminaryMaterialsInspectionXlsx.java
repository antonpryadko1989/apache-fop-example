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
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.services.Const.*;
import static com.visoft.services.POIXls.*;

public class PreliminaryMaterialsInspectionXlsx implements XLSXBuilder {

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
        FontParams fontCalibri12 = setFontParams(FONT_NAME_CALIBRI, false, (short) (12 * FONT_RATE));

        XSSFCellStyle lTMediumRBThinCalibri12Bold = getLTMediumRBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tMediumLRBThinCalibri12Bold = getTMediumLRBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tRMediumLBThinCalibri12Bold = getTRMediumLBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle bMediumLTRThinCalibri12Bold = getBMediumLTRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lBMediumRTThinCalibri12Bold = getLBMediumRTThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle rBMediumLTThinCalibri12Bold = getRBMediumLTThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lMediumTRBThinCalibri12Bold = getLMediumTRBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle allBordersThinCalibri12Bold = setAllBordersThin(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lBMediumTRThinCalibri12Bold = getLBMediumTRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lTBMediumRThinCalibri12Bold = getLTBMediumRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tBMediumLRThinCalibri12Bold = getTBMediumLRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tRBMediumLThinCalibri12Bold = getTRBMediumLThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tRBMediumLThinCalibri12bold = getTRBMediumLThinFont(xlsxObject, fontCalibri12bold);

        XSSFCellStyle tBMediumLRThinCalibri12 = getTBMediumLRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tRBMediumLThinCalibri12 = getTRBMediumLThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle rBMediumLTThinCalibri12 = getRBMediumLTThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tMediumLRBThinCalibri12 = getTMediumLRBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tRMediumLBThinCalibri12 = getTRMediumLBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle bMediumLTRThinCalibri12 = getBMediumLTRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle rMediumLTBThinCalibri12 = getRMediumLTBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle allBordersThinCalibri12 = setAllBordersThin(xlsxObject, fontCalibri12);
        XSSFCellStyle borderMediumCalibri20BoldAligmentRight = setAlignmentByParam(xlsxObject, fontCalibri20bold, Alignment.RIGHT, BorderStyle.MEDIUM);

        addLogo(xlsxObject, template.getBody());
        addHeaderInfo(xlsxObject, PRELIM_MATERIALS_INS_TEMPLATE_NAME, VERSION, DATE, template.getBody().getBodyElements().get(TEMPLATE_NAME_VAL),
                template.getBody().getBodyElements().get(VERSION_VAL), template.getBody().getBodyElements().get(DATE_VAL),
                lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12Bold, tRMediumLBThinCalibri12Bold,
                lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12Bold, rBMediumLTThinCalibri12Bold);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(xlsxObject,template,
                lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tMediumLRBThinCalibri12Bold, tRMediumLBThinCalibri12,
                lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, allBordersThinCalibri12Bold, rMediumLTBThinCalibri12,
                lBMediumTRThinCalibri12Bold, bMediumLTRThinCalibri12, bMediumLTRThinCalibri12Bold, rBMediumLTThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addPreliminaryMaterialsInspectionBody(xlsxObject, template.getBody().getBodyElements(),
                lTMediumRBThinCalibri12Bold, tRMediumLBThinCalibri12, lMediumTRBThinCalibri12Bold,
                rMediumLTBThinCalibri12, lBMediumRTThinCalibri12Bold, rBMediumLTThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, PRELIMINARY_INSPECTION_RESULTS,
                borderMediumCalibri20BoldAligmentRight);
        addPreliminaryInspectionResults(xlsxObject,
                PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION, PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS,
                PRELIM_INSPEC_RESULT_INSPECTION_RESULTS, PRELIM_INSPEC_RESULT_CERTIFICATE_NO,
                PRELIM_INSPEC_RESULT_PASS_FAIL, PRELIM_INSPEC_RESULT_COMMENTS,
                lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12Bold);
        if (notEmpty) {
            addPreliminaryInspectionResultsList(xlsxObject, template.getBody().getBodyLists().get(PRELIMINARY_INSPECTION_RESULTS_VAL),
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                    lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
        } else {
            addPreliminaryInspectionResultsList(xlsxObject, null,
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                    lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, ADDITIONAL_DOCUMENTS, borderMediumCalibri20BoldAligmentRight);

        addAdditionalDocuments(xlsxObject, ADDITIONAL_DOCUMENTS_ITEM,
                ADDITIONAL_DOCUMENTS_EXISTS, ADDITIONAL_DOCUMENTS_CERTIFICATE_NO,
                ADDITIONAL_DOCUMENTS_EXPIRATION, ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS,
                lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold);
        if (notEmpty) {
            addAdditionalDocumentsList(xlsxObject, template.getBody().getBodyLists().get(ADDITIONAL_DOCUMENTS_VAL),
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                    lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
        } else {
            addAdditionalDocumentsList(xlsxObject, null,
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                    lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, APPROVALS, borderMediumCalibri20BoldAligmentRight);
        addApprovals(xlsxObject, APPROVALS_ROLE, APPROVALS_NAME, APPROVALS_SIGNATURE, APPROVALS_DATE, APPROVALS_STATUS,
                lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold);
        if (notEmpty) {
            addApprovalsList(xlsxObject, template.getBody().getBodyLists().get(APPROVALS_VAL),
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12);
        } else {
            addApprovalsList(xlsxObject, null,
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12);
        }
        return input2OutputService.writeWorkBook(workbook);
    }

    private void addPreliminaryMaterialsInspectionBody(XLSXObject xlsxObject, Map<String,String> currentMap,
                                                       XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                                                       XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6) {
        setValuesToRow(xlsxObject, getPreliminaryMaterialsInspectionBodyRows(currentMap),
                style1, style2, style3, style4, style5, style6);
    }

    private static LinkedMap<String,String> getPreliminaryMaterialsInspectionBodyRows(
            Map<String, String> body) {
        LinkedMap<String, String> result = new LinkedMap<>();
        result.put(PRELIM_MATERIALS_INS_APPROVAL_NO, body
                .get(PRELIM_MATERIALS_INS_APPROVAL_NO_VAL));
        result.put(PRELIM_MATERIALS_INS_SUPPLIER_NAME, body
                .get(PRELIM_MATERIALS_INS_SUPPLIER_NAME_VAL));
        result.put(PRELIM_MATERIALS_INS_SUBPROJECT, body
                .get(PRELIM_MATERIALS_INS_SUBPROJECT_VAL));
        result.put(PRELIM_MATERIALS_INS_SOURCE_OF_MATERIAL, body
                .get(PRELIM_MATERIALS_INS_SOURCE_OF_MATERIAL_VAL));
        result.put(PRELIM_MATERIALS_INS_MATERIAL_SUPPLIED, body
                .get(PRELIM_MATERIALS_INS_MATERIAL_SUPPLIED_VAL));
        result.put(PRELIM_MATERIALS_INS_MATERIAL_USE, body
                .get(PRELIM_MATERIALS_INS_MATERIAL_USE_VAL));
        return result;
    }
}