package com.visoft.services.impl;

import com.visoft.cellStyleUtil.FontParams;
import com.visoft.services.Alignment;
import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import com.visoft.templates.enums.ConfigType;
import org.apache.commons.collections4.map.LinkedMap;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.CellStyleUtil.getTRBMediumLThinFont;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.services.Const.*;
import static com.visoft.services.POIXls.*;

public class NCRXlsx implements XLSXBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(TemplateDTO template) throws IOException {
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());

        FontParams fontCalibri20Bold = setFontParams(FONT_NAME_CALIBRI, true, (short) (20 * FONT_RATE));
        FontParams fontCalibri12bold = setFontParams(FONT_NAME_CALIBRI, true, (short) (12 * FONT_RATE));
        FontParams fontCalibri12 = setFontParams(FONT_NAME_CALIBRI, false, (short) (12 * FONT_RATE));

        XSSFCellStyle borderMediumCalibri20BoldAligmentCenter = setAlignmentByParam(xlsxObject, fontCalibri20Bold, Alignment.CENTER, BorderStyle.MEDIUM);
        XSSFCellStyle borderMediumCalibri20BoldAligmentRight = setAlignmentByParam(xlsxObject, fontCalibri20Bold, Alignment.RIGHT, BorderStyle.MEDIUM);

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
        XSSFCellStyle lMediumTRBThinCalibri12 = getLMediumTRBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tBMediumLRThinCalibri12 = getTBMediumLRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tRBMediumLThinCalibri12 = getTRBMediumLThinFont(xlsxObject, fontCalibri12);

        addLogo(xlsxObject, template.getBody());
        addHeaderInfo(xlsxObject, NCR_TEMPLATE_NAME, VERSION, DATE,
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
        setNCRNumber(xlsxObject, NCR_NUMBER, template.getBody().getBodyElements().get(NCR_NUMBER_VAL),
                lTBMediumRThinCalibri12Bold, tRBMediumLThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addQaQc(xlsxObject, NCR_OPENED, DATE_OF_NCR,
                template.getBody().getBodyElements().get(QC_OPENED_NAME_VAL),
                template.getBody().getBodyElements().get(QC_OPENED_DATE_OF_NCR_VAL),
                template.getBody().getBodyElements().get(QA_OPENED_NAME_VAL),
                template.getBody().getBodyElements().get(QA_OPENED_DATE_OF_NCR_VAL),
                lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold,
                tMediumLRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                bMediumLTRThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12);
        addAlignmentRow(xlsxObject, DETAILS_OF_STRUCTURE, borderMediumCalibri20BoldAligmentCenter);
        addNCRBody(xlsxObject, template,
                lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12Bold, tRMediumLBThinCalibri12Bold,
                lMediumTRBThinCalibri12, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                lMediumTRBThinCalibri12Bold, lBMediumRTThinCalibri12Bold, rBMediumLTThinCalibri12);
        addAlignmentRow(xlsxObject, ACCEPTANCE_OF_CORRECTIVE_ACTION, borderMediumCalibri20BoldAligmentCenter);
        addQaQc(xlsxObject, NCR_CLOSED, CLOSING_DATE,
                template.getBody().getBodyElements().get(QC_CLOSED_NAME_VAL),
                template.getBody().getBodyElements().get(QC_CLOSED_DATE_OF_NCR_VAL),
                template.getBody().getBodyElements().get(QA_CLOSED_NAME_VAL),
                template.getBody().getBodyElements().get(QA_CLOSED_DATE_OF_NCR_VAL),
                lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold,
                tMediumLRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                bMediumLTRThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, ADDITIONAL_DOCUMENTS,
                    borderMediumCalibri20BoldAligmentRight);
        addAdditionalDocuments(xlsxObject, ADDITIONAL_DOCUMENTS_ITEM, ADDITIONAL_DOCUMENTS_EXISTS,
                ADDITIONAL_DOCUMENTS_CERTIFICATE_NO, ADDITIONAL_DOCUMENTS_EXPIRATION, ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS,
                lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold);
        if(notEmpty) {
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
        return input2OutputService.writeWorkBook(workbook);
    }

    private void addNCRBody(XLSXObject xlsxObject, TemplateDTO template,
                            XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3,
                            XSSFCellStyle style4, XSSFCellStyle style5, XSSFCellStyle style6,
                            XSSFCellStyle style7,  XSSFCellStyle style8, XSSFCellStyle style9) {
        ConfigType configType = template.getConfigType();
        switch (configType) {
            case ZERO:
                setValuesToRow(xlsxObject, configType, NCR_ZERO_ROW_HEADER, style1, style2, style3);
                break;
            case ONE:
                setValuesToRow(xlsxObject, configType, NCR_ONE_ROW_HEADER, style1, style2, style3);
                break;
            default:
                setValuesToRow(xlsxObject, configType, NCR_TWO_ROW_HEADER, style1, style2, style3);
                break;
        }
        setValuesToRow(xlsxObject, configType, rowArrayNCRBody(template), style4, style5, style6);
        setValuesToRow(xlsxObject, getNCRBodyRows(template.getBody().getBodyElements(), configType), style7, style6, style8, style9);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
    }

    private String[] rowArrayNCRBody(TemplateDTO template) {
        switch (template.getConfigType()){
            case ZERO:
                return new String[]{
                        template.getBody().getBodyElements().get(ELEMENT_STATION_ROAD_TUNNEL_BRIDGE_VAL),
                        template.getBody().getBodyElements().get(STRUCTURE_VAL),
                        template.getBody().getBodyElements().get(ELEMENT_VAL),
                        template.getBody().getBodyElements().get(FROM_SECTION_VAL),
                        template.getBody().getBodyElements().get(TILL_SECTION_VAL),
                        template.getBody().getBodyElements().get(SIDE_VAL),
                        template.getBody().getBodyElements().get(NCR_LEVEL_VAL),
                        template.getBody().getBodyElements().get(NUMBER_OF_DAYS_LATE_VAL)
                };
            case ONE:
                return new String[]{
                        template.getBody().getBodyElements().get(STRUCTURE_VAL),
                        template.getBody().getBodyElements().get(ELEMENT_LAYER_VAL),
                        template.getBody().getBodyElements().get(FROM_SECTION_VAL),
                        template.getBody().getBodyElements().get(TILL_SECTION_VAL),
                        template.getBody().getBodyElements().get(SIDE_VAL),
                        template.getBody().getBodyElements().get(NCR_LEVEL_VAL),
                        template.getBody().getBodyElements().get(BLUE_BOOK_CODE_VAL)
                };
            default:
                return new String[]{
                        template.getBody().getBodyElements().get(ELEMENT_STATION_ROAD_TUNNEL_BRIDGE_VAL),
                        template.getBody().getBodyElements().get(STRUCTURE_VAL),
                        template.getBody().getBodyElements().get(ELEMENT_LAYER_VAL),
                        template.getBody().getBodyElements().get(SUB_ELEMENT_VAL),
                        template.getBody().getBodyElements().get(FROM_SECTION_VAL),
                        template.getBody().getBodyElements().get(TILL_SECTION_VAL),
                        template.getBody().getBodyElements().get(SIDE_VAL),
                        template.getBody().getBodyElements().get(NCR_LEVEL_VAL)
                };
        }
    }

    private LinkedMap<String, String> getNCRBodyRows(Map<String, String> body, ConfigType configType) {
        switch (configType){
            case ZERO:
                return getConfigZeroNCRBodyRows(body);
            case ONE:
                return getConfigOneNCRBodyRows(body);
            default:
                return getConfigTwoNCRBodyRows(body);
        }
    }

    private LinkedMap<String,String> getConfigZeroNCRBodyRows(Map<String, String> body ) {
        LinkedMap<String, String> result = new LinkedMap<>();
        result.put(EXPECTED_CLOSING_DATE, body.get(EXPECTED_CLOSING_DATE_VAL));
        result.put(UPDATED_EXPECTED_CLOSING_DATE, body.get(UPDATED_EXPECTED_CLOSING_DATE_VAL));
        result.put(SUB_PROJECT, body.get(SUB_PROJECT_VAL));
        result.put(NCR_DESCRIPTION, body.get(NCR_DESCRIPTION_VAL));
        result.put(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER, body.get(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER_VAL));
        result.put(CORRECTIVE_ACTION, body.get(CORRECTIVE_ACTION_VAL));
        result.put(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION, body.get(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL));
        result.put(RESPONSIBLE_PERSON, body.get(RESPONSIBLE_PERSON_VAL));
        result.put(REMARKS, body.get(REMARKS_VAL));
        return result;
    }

    private static LinkedMap<String,String> getConfigOneNCRBodyRows (Map<String, String> body) {
        LinkedMap<String, String> result = new LinkedMap<>();
        result.put(APPROXIMATE_CLOSING_DATE_NCR_ONE, body.get(APPROXIMATE_CLOSING_DATE_VAL));
        result.put(NUMBER_OF_DAYS_LATE, body.get(NUMBER_OF_DAYS_LATE_VAL));
        result.put(NCR_DESCRIPTION, body.get(NCR_DESCRIPTION_VAL));
        result.put(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER, body.get(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER_VAL));
        result.put(RESPONSIBLE_PARTY_ROLE, body.get(RESPONSIBLE_PARTY_ROLE_VAL));
        result.put(IS_OPENED_BEFORE, body.get(IS_OPENED_BEFORE_VAL));
        result.put(OPENED_BEFORE_NCR_NO, body.get(OPENED_BEFORE_NCR_NO_VAL));
        result.put(CORRECTIVE_ACTION, body.get(CORRECTIVE_ACTION_VAL));
        result.put(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION, body.get(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL));
        result.put(TYPE_OF_TEST, body.get(TYPE_OF_TEST_VAL));
        result.put(REMARKS, body.get(REMARKS_VAL));
        return result;
    }

    private LinkedMap<String,String> getConfigTwoNCRBodyRows(Map<String, String> body) {
        LinkedMap<String, String> result = new LinkedMap<>();
        result.put(APPROXIMATE_CLOSING_DATE_NCR_TWO, body.get(APPROXIMATE_CLOSING_DATE_VAL));
        result.put(ORIGINAL_EXPECTED_CLOSING_DATE, body.get(ORIGINAL_EXPECTED_CLOSING_DATE_VAL));
        result.put(NUMBER_OF_DAYS_LATE, body.get(NUMBER_OF_DAYS_LATE_VAL));
        result.put(DAMAGE, body.get(DAMAGE_VAL));
        result.put(QUALITY_IMPACT, body.get(QUALITY_IMPACT_VAL));
        result.put(NCR_DESCRIPTION, body.get(NCR_DESCRIPTION_VAL));
        result.put(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER, body.get(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER_VAL));
        result.put(CORRECTIVE_ACTION, body.get(CORRECTIVE_ACTION_VAL));
        result.put(RESPONSIBLE_PERSON, body.get(RESPONSIBLE_PERSON_VAL));
        result.put(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION, body.get(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL));
        result.put(REMARKS, body.get(REMARKS_VAL));
        return result;
    }
}