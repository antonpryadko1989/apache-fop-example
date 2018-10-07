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
import static com.visoft.cellStyleUtil.CellStyleUtil.getTRBMediumLThinFont;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.services.Const.*;
import static com.visoft.services.POIXls.*;
public class POCXlsx implements XLSXBuilder{

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
        XSSFCellStyle lTMediumRBThinCalibri12 = getLTMediumRBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle lMediumTRBThinCalibri12 = getLMediumTRBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle lBMediumRTThinCalibri12 = getLBMediumRTThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle lTBMediumRThinCalibri12 = getLTBMediumRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tBMediumLRThinCalibri12 = getTBMediumLRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tRBMediumLThinCalibri12 = getTRBMediumLThinFont(xlsxObject, fontCalibri12);

        addLogo(xlsxObject, template.getBody());
        addHeaderInfo(xlsxObject, POC_TEMPLATE_NAME, VERSION, DATE,
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
        if(notEmpty) {
            if (template.getBody().getBodyLists().get(DRAWINGS_VAL) != null) {
                addDrawings(xlsxObject, DRAWING_NO, DRAWING_VERSION_REVISION, DRAWING_NAME,
                        lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold);
                addDrawingsList(xlsxObject, template.getBody().getBodyLists().get(DRAWINGS_VAL),
                        lTMediumRBThinCalibri12, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                        lMediumTRBThinCalibri12, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                        lBMediumRTThinCalibri12, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                        lTBMediumRThinCalibri12, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
                addEmptyStringWithGivenHeight(xlsxObject, 9);
            }
        }
        addPOCBody(xlsxObject, template.getBody().getBodyElements(),
                lTMediumRBThinCalibri12Bold, tRMediumLBThinCalibri12,
                lMediumTRBThinCalibri12Bold, rMediumLTBThinCalibri12,
                lBMediumRTThinCalibri12Bold, rBMediumLTThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, ADDITIONAL_DOCUMENTS, borderMediumCalibri20BoldAligmentRight);
        addAdditionalDocuments(xlsxObject, ADDITIONAL_DOCUMENTS_ITEM, ADDITIONAL_DOCUMENTS_EXISTS, ADDITIONAL_DOCUMENTS_CERTIFICATE_NO,
                ADDITIONAL_DOCUMENTS_EXPIRATION, ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS,
                lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold);
        if(notEmpty){
            addAdditionalDocumentsList(xlsxObject, template.getBody().getBodyLists().get(ADDITIONAL_DOCUMENTS_VAL),
                    lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                    lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                    lBMediumRTThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                    lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
        }else{
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

    private void addPOCBody(XLSXObject xlsxObject,
                            Map<String, String> currentMap,
                            XSSFCellStyle style1, XSSFCellStyle style2,
                            XSSFCellStyle style3, XSSFCellStyle style4,
                            XSSFCellStyle style5, XSSFCellStyle style6){
        setValuesToRow(xlsxObject, getPOCBodyRows(currentMap), style1, style2, style3, style4, style5, style6);
    }

    private LinkedMap<String,String> getPOCBodyRows(Map<String, String> body) {
        LinkedMap<String, String> result = new LinkedMap<>();
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