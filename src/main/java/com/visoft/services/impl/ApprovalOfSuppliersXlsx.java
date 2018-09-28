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

public class ApprovalOfSuppliersXlsx implements XLSXBuilder{
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
        addHeaderInfo(xlsxObject, APP_OF_SUPPL_TEMPLATE_NAME, VERSION, DATE,
                template.getBody().getBodyElements().get(TEMPLATE_NAME_VAL),
                template.getBody().getBodyElements().get(VERSION_VAL),
                template.getBody().getBodyElements().get(DATE_VAL),
                borderThinCalibri18Bold, borderThinCalibri12Bold);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(xlsxObject, template,
                borderThinCalibri12Bold, borderThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addApprovalOfSuppliersBody(xlsxObject,
                template.getBody().getBodyElements(),
                borderThinCalibri12Bold, borderThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, CERTIFICATIONS,
                borderThinCalibri20BoldAligmentRight);
        addCertifications(xlsxObject, CERTIFICATIONS_ITEM, CERTIFICATIONS_EXISTS,
                CERTIFICATIONS_CERTIFICATE_NO, CERTIFICATIONS_EXPIRATION,
                CERTIFICATIONS_ATTACHED_DOCUMENTS, borderThinCalibri12Bold);
        if (notEmpty) {
            addCertificationsList(xlsxObject, template.getBody().getBodyLists()
                            .get(CERTIFICATIONS_VAL),
                    borderThinCalibri12Bold, borderThinCalibri12);
        } else {
            addCertificationsList(xlsxObject, null,
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
        System.out.println(APPROVAL_OF_SUPPLIERS+ " " +
                (System.currentTimeMillis() - start));
        return writeWorkBook(workbook);
    }

    private void addApprovalOfSuppliersBody(XLSXObject xlsxObject,
                                            Map<String, String>
                                                          currentMap,
                                            XSSFCellStyle styleOne,
                                            XSSFCellStyle styleTwo){
        setValuesToRow(xlsxObject, getApprovalOfSuppliersBodyRows(currentMap),
                styleOne, styleTwo);
    }

    private static Map<String,String> getApprovalOfSuppliersBodyRows(
            Map<String, String> body) {
        Map<String, String> result = new LinkedHashMap<>();
        result.put(APP_OF_SUPPL_APPROVAL_NO, body
                .get(APP_OF_SUPPL_APPROVAL_NO_VAL));
        result.put(APP_OF_SUPPL_SUPPLIER_NAME, body
                .get(APP_OF_SUPPL_SUPPLIER_NAME_VAL));
        result.put(APP_OF_SUPPL_SUBPROJECT, body
                .get(APP_OF_SUPPL_SUBPROJECT_VAL));
        result.put(APP_OF_SUPPL_CONTACT, body
                .get(APP_OF_SUPPL_CONTACT_VAL));
        result.put(APP_OF_SUPPL_SUPPLIED_MATERIALS, body
                .get(APP_OF_SUPPL_SUPPLIED_MATERIALS_VAL));
        return result;
    }
}
