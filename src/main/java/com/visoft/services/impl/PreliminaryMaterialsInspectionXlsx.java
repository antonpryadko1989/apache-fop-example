package com.visoft.services.impl;

import com.visoft.services.Alignment;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.xssf.usermodel.*;

import java.util.*;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.services.Const.*;
import static com.visoft.services.POIXls.*;

public class PreliminaryMaterialsInspectionXlsx implements XLSXBuilder {
    @Override
    public String buildXLSX(TemplateDTO template) {
        long start = System.currentTimeMillis();
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XLSXObject xlsxObject = new XLSXObject(startRowNum,
                workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());
        XSSFCellStyle borderThinCalibri20BoldAligmentRight = setAlignmentByParam(xlsxObject,
                setFontParams(FONT_NAME_CALIBRI, true,
                        (short) (20 * FONT_RATE)), Alignment.RIGHT);
        XSSFCellStyle borderThinCalibri18Bold = setAllBordersThin(xlsxObject,
                setFontParams(FONT_NAME_CALIBRI, true,
                        (short) (18 * FONT_RATE)));
        XSSFCellStyle borderThinCalibri12Bold = setAllBordersThin(xlsxObject,
                setFontParams(FONT_NAME_CALIBRI, true,
                        (short) (12 * FONT_RATE)));
        XSSFCellStyle borderThinCalibri12 = setAllBordersThin(xlsxObject,
                setFontParams(FONT_NAME_CALIBRI, false,
                        (short) (12 * FONT_RATE)));
        addLogo(xlsxObject,
                template.getBody().getBodyElements().get(LOGO_PATH),
                template.getBody().getBodyElements().get(LOGO_VAL));
        addHeaderInfo(xlsxObject, PRELIM_MATERIALS_INS_TEMPLATE_NAME, VERSION, DATE,
                template.getBody().getBodyElements().get(TEMPLATE_NAME_VAL),
                template.getBody().getBodyElements().get(VERSION_VAL),
                template.getBody().getBodyElements().get(DATE_VAL),
                borderThinCalibri18Bold, borderThinCalibri12Bold);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(xlsxObject,template,
                borderThinCalibri12Bold, borderThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addPreliminaryMaterialsInspectionBody(xlsxObject,
                template.getBody().getBodyElements(),
                borderThinCalibri12Bold,
                borderThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, PRELIMINARY_INSPECTION_RESULTS,
                borderThinCalibri20BoldAligmentRight);
        setValuesToRow(xlsxObject,
                PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION,
                PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS,
                PRELIM_INSPEC_RESULT_INSPECTION_RESULTS,
                PRELIM_INSPEC_RESULT_CERTIFICATE_NO,
                PRELIM_INSPEC_RESULT_PASS_FAIL,
                PRELIM_INSPEC_RESULT_COMMENTS,
                borderThinCalibri12Bold,
                true);
        if (notEmpty) {
            addPrelimitaryInspectionResultsList(xlsxObject, template.getBody()
                            .getBodyLists().get(PRELIMINARY_INSPECTION_RESULTS_VAL),
                    borderThinCalibri12);
        } else {
            addPrelimitaryInspectionResultsList(xlsxObject, null,
                    borderThinCalibri12);
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, ADDITIONAL_DOCUMENTS,
                borderThinCalibri20BoldAligmentRight);
        addAdditionalDocuments(xlsxObject, ADDITIONAL_DOCUMENTS_ITEM,
                ADDITIONAL_DOCUMENTS_EXISTS,
                ADDITIONAL_DOCUMENTS_CERTIFICATE_NO,
                ADDITIONAL_DOCUMENTS_EXPIRATION,
                ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS,
                borderThinCalibri12Bold);
        if (notEmpty) {
            addApprovalsList(xlsxObject, template.getBody().getBodyLists()
                            .get(ADDITIONAL_DOCUMENTS_VAL),
                    borderThinCalibri12Bold, borderThinCalibri12);
        } else {
            addApprovalsList(xlsxObject, null,
                borderThinCalibri12Bold, borderThinCalibri12);
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, APPROVALS,
                borderThinCalibri20BoldAligmentRight);
        addApprovals(xlsxObject, APPROVALS_ROLE,
                APPROVALS_NAME, APPROVALS_SIGNATURE,
                APPROVALS_DATE, APPROVALS_STATUS,
                borderThinCalibri12Bold);
        if (notEmpty) {
            addApprovalsList(xlsxObject, template.getBody().getBodyLists()
                            .get(APPROVALS_VAL),
                    borderThinCalibri12Bold, borderThinCalibri12);
        } else {
            addApprovalsList(xlsxObject, null,
                    borderThinCalibri12Bold, borderThinCalibri12);
        }
        System.out.println(PRELIMINARY_MATERIALS_INSPECTION + " " +
                (System.currentTimeMillis() - start));
        return writeWorkBook(workbook);
    }

    private void
    addPreliminaryMaterialsInspectionBody(XLSXObject xlsxObject,
                                          Map<String,String> currentMap,
                                          XSSFCellStyle styleOne,
                                          XSSFCellStyle styleTwo) {
        setValuesToRow(xlsxObject,
                getPreliminaryMaterialsInspectionBodyRows(currentMap),
                styleOne, styleTwo);
    }

    private static Map<String,String> getPreliminaryMaterialsInspectionBodyRows(
            Map<String, String> body) {
        Map<String, String> result = new LinkedHashMap<>();
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
