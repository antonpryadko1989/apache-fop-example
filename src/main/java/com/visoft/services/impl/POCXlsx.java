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
public class POCXlsx implements XLSXBuilder{

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
        addHeaderInfo(xlsxObject,
                POC_TEMPLATE_NAME,
                VERSION,
                DATE,
                template.getBody().getBodyElements().get(TEMPLATE_NAME_VAL),
                template.getBody().getBodyElements().get(VERSION_VAL),
                template.getBody().getBodyElements().get(DATE_VAL),
                borderThinCalibri18Bold,
                borderThinCalibri12Bold);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(xlsxObject, template,
                borderThinCalibri12Bold, borderThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);

        addDrawings(xlsxObject, DRAWING_NO, DRAWING_VERSION_REVISION,
                DRAWING_NAME, borderThinCalibri12Bold, true);
        if(notEmpty){
            addDrawingsList(xlsxObject,
                    template.getBody().getBodyLists().get(DRAWINGS_VAL),
                    borderThinCalibri12);
        } else {
            addDrawingsList(xlsxObject, null,
                    borderThinCalibri12);
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addPOCBody(xlsxObject, template.getBody()
                .getBodyElements(), borderThinCalibri12Bold,
                borderThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, ADDITIONAL_DOCUMENTS,
                borderThinCalibri20BoldAligmentRight);
        addAdditionalDocuments(xlsxObject, ADDITIONAL_DOCUMENTS_ITEM,
                ADDITIONAL_DOCUMENTS_EXISTS,
                ADDITIONAL_DOCUMENTS_CERTIFICATE_NO,
                ADDITIONAL_DOCUMENTS_EXPIRATION,
                ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS,
                borderThinCalibri12Bold);
        if(notEmpty){
            addAdditionalDocumentsList(xlsxObject, template.getBody().getBodyLists()
                            .get(ADDITIONAL_DOCUMENTS_VAL),
                    borderThinCalibri12Bold, borderThinCalibri12);
        }else{
            addAdditionalDocumentsList(xlsxObject, null,
                    borderThinCalibri12Bold, borderThinCalibri12);
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, APPROVALS,
                borderThinCalibri20BoldAligmentRight);
        if(notEmpty){
            addApprovals(xlsxObject, APPROVALS_ROLE,
                    APPROVALS_NAME, APPROVALS_SIGNATURE,
                    APPROVALS_DATE, APPROVALS_STATUS,
                    borderThinCalibri12Bold);
        }else {
            addApprovalsList(xlsxObject, null,
                    borderThinCalibri12Bold, borderThinCalibri12);
        }
        System.out.println(POC + " " +
                (System.currentTimeMillis() - start));
        return writeWorkBook(workbook);
    }

    private void addPOCBody(XLSXObject xlsxObject,
                            Map<String, String> currentMap,
                            XSSFCellStyle styleOne,
                            XSSFCellStyle styleTwo){
        setValuesToRow(xlsxObject, getPOCBodyRows(currentMap),
                styleOne, styleTwo);
    }

    private Map<String,String> getPOCBodyRows(Map<String, String> body) {
        Map<String, String> result = new LinkedHashMap<>();
        result.put(POC_REPORT_NO, body.get(POC_REPORT_NO_VAL));
        result.put(POC_TYPE_OF_TEST, body.get(POC_TYPE_OF_TEST_VAL));
        result.put(POC_ELEMENT, body.get(POC_TYPE_OF_TEST_VAL));
        result.put(POC_FROM_SECTION, body.get(POC_FROM_SECTION_VAL));
        result.put(POC_UP_TO_SECTION, body.get(POC_UP_TO_SECTION_VAL));
        result.put(POC_SIDE, body.get(POC_SIDE_VAL));
        result.put(POC_PARTICIPANTS_IN_TEST, body.get(POC_PARTICIPANTS_IN_TEST_VAL));
        result.put(POC_MATERIAL_FOR_USE, body.get(POC_MATERIAL_FOR_USE_VAL));
        result.put(POC_TOOLS_USED, body.get(POC_TOOLS_USED_VAL));
        result.put(POC_DATE_OF_EXECUTION, body.get(POC_DATE_OF_EXECUTION_VAL));
        result.put(POC_DESCRIPTION_OF_TEST, body.get(POC_DESCRIPTION_OF_TEST_VAL));
        result.put(POC_CONCLUSIONS_OF_TEST, body.get(POC_CONCLUSIONS_OF_TEST_VAL));
        result.put(POC_CORRECTIVE_ACTION, body.get(POC_CORRECTIVE_ACTION_VAL));
        return result;
    }


}
