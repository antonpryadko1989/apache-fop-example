package com.visoft.services.impl;

import com.visoft.services.Alignment;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import com.visoft.templates.enums.ConfigType;
import org.apache.poi.xssf.usermodel.*;

import java.util.*;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.services.Const.*;
import static com.visoft.services.POIXls.*;

public class NCRXlsx implements XLSXBuilder {
    @Override
    public String buildXLSX(TemplateDTO template) {
        long start = System.currentTimeMillis();
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XLSXObject xlsxObject = new XLSXObject(startRowNum,
                workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());
        XSSFCellStyle borderThinCalibri20BoldAligmentCenter = setAlignmentByParam(xlsxObject,
                setFontParams(FONT_NAME_CALIBRI, true,
                        (short) (20 * FONT_RATE)), Alignment.CENTER);
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
                NCR_TEMPLATE_NAME,
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
        setNCRNumber(xlsxObject, NCR_NUMBER, template.getBody()
                        .getBodyElements().get(NCR_NUMBER_VAL),
                borderThinCalibri12Bold, borderThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addQaQc(xlsxObject, NCR_OPENED, DATE_OF_NCR,
                template.getBody().getBodyElements().get(QC_OPENED_NAME_VAL),
                template.getBody().getBodyElements().get(QC_OPENED_DATE_OF_NCR_VAL),
                template.getBody().getBodyElements().get(QA_OPENED_NAME_VAL),
                template.getBody().getBodyElements().get(QA_OPENED_DATE_OF_NCR_VAL),
                borderThinCalibri12Bold, borderThinCalibri12);
        addAlignmentRow(xlsxObject, DETAILS_OF_STRUCTURE,
                borderThinCalibri20BoldAligmentCenter);
        addNCRBody(xlsxObject, template,
                borderThinCalibri12Bold, borderThinCalibri12);
        addAlignmentRow(xlsxObject, ACCEPTANCE_OF_CORRECTIVE_ACTION,
                borderThinCalibri20BoldAligmentCenter);
        addQaQc(xlsxObject, NCR_CLOSED, CLOSING_DATE,
                template.getBody().getBodyElements().get(QC_CLOSED_NAME_VAL),
                template.getBody().getBodyElements().get(QC_CLOSED_DATE_OF_NCR_VAL),
                template.getBody().getBodyElements().get(QA_CLOSED_NAME_VAL),
                template.getBody().getBodyElements().get(QA_CLOSED_DATE_OF_NCR_VAL),
                borderThinCalibri12Bold, borderThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(xlsxObject, ADDITIONAL_DOCUMENTS,
                    borderThinCalibri20BoldAligmentRight);
            addAdditionalDocuments(xlsxObject, ADDITIONAL_DOCUMENTS_ITEM,
                    ADDITIONAL_DOCUMENTS_EXISTS,
                    ADDITIONAL_DOCUMENTS_CERTIFICATE_NO,
                    ADDITIONAL_DOCUMENTS_EXPIRATION,
                    ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS,
                    borderThinCalibri12Bold);
        if(notEmpty) {
            addAdditionalDocumentsList(xlsxObject, template.getBody().getBodyLists()
                            .get(ADDITIONAL_DOCUMENTS_VAL),
                    borderThinCalibri12Bold, borderThinCalibri12);
        } else {
            addAdditionalDocumentsList(xlsxObject, new ArrayList<>(),
                    borderThinCalibri12Bold, borderThinCalibri12);
        }
        System.out.println(NCR + " " +
                (System.currentTimeMillis() - start));
        return writeWorkBook(workbook);
    }

    private void addNCRBody(XLSXObject xlsxObject,
                            TemplateDTO template,
                            XSSFCellStyle styleOne,
                            XSSFCellStyle styleTwo) {
        ConfigType configType = template.getConfigType();
        switch (configType){
            case ZERO:
                setValuesToRow(xlsxObject, configType, NCR_ZERO_ROW_HEADER,
                        styleOne, true);
                break;
            case ONE:
                setValuesToRow(xlsxObject, configType, NCR_ONE_ROW_HEADER,
                        styleOne, true);
                break;
            case TWO:
                setValuesToRow(xlsxObject, configType, NCR_TWO_ROW_HEADER,
                        styleOne, true);
                break;
            default:
                break;
        }
        setValuesToRow(xlsxObject, configType, rowArrayNCRBody(template),
                styleTwo, false);
        setValuesToRow(xlsxObject, getNCRBodyRows(template.getBody()
                        .getBodyElements(),
                configType),
                styleOne, styleTwo);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
    }

    private String[] rowArrayNCRBody(TemplateDTO template) {
        switch (template.getConfigType()){
            case ZERO:
                return new String[]{
                        template.getBody().getBodyElements()
                                .get(ELEMENT_STATION_ROAD_TUNNEL_BRIDGE_VAL),
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
            case TWO:
                return new String[]{
                        template.getBody().getBodyElements()
                                .get(ELEMENT_STATION_ROAD_TUNNEL_BRIDGE_VAL),
                        template.getBody().getBodyElements().get(STRUCTURE_VAL),
                        template.getBody().getBodyElements().get(ELEMENT_LAYER_VAL),
                        template.getBody().getBodyElements().get(SUB_ELEMENT_VAL),
                        template.getBody().getBodyElements().get(FROM_SECTION_VAL),
                        template.getBody().getBodyElements().get(TILL_SECTION_VAL),
                        template.getBody().getBodyElements().get(SIDE_VAL),
                        template.getBody().getBodyElements().get(NCR_LEVEL_VAL)

                };
            default:
                return new String[0];
        }
    }

    private Map<String, String> getNCRBodyRows(
            Map<String, String> body, ConfigType configType) {
        switch (configType){
            case ZERO:
                return getConfigZeroNCRBodyRows(body);
            case ONE:
                return getConfigOneNCRBodyRows(body);
            default:
                return getConfigTwoNCRBodyRows(body);
        }
    }

    private Map<String,String> getConfigZeroNCRBodyRows(
            Map<String, String> body ) {
        Map<String, String> result = new LinkedHashMap<>();
        result.put(EXPECTED_CLOSING_DATE, body.get(EXPECTED_CLOSING_DATE_VAL));
        result.put(UPDATED_EXPECTED_CLOSING_DATE, body.get(UPDATED_EXPECTED_CLOSING_DATE_VAL));
        result.put(SUB_PROJECT, body.get(SUB_PROJECT_VAL));
        result.put(NCR_DESCRIPTION, body.get(NCR_DESCRIPTION_VAL));
        result.put(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER,
                body.get(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER_VAL));
        result.put(CORRECTIVE_ACTION, body.get(CORRECTIVE_ACTION_VAL));
        result.put(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION,
                body.get(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL));
        result.put(RESPONSIBLE_PERSON, body.get(RESPONSIBLE_PERSON_VAL));
        result.put(REMARKS, body.get(REMARKS_VAL));
        return result;
    }

    private static Map<String,String> getConfigOneNCRBodyRows (
            Map<String, String> body) {
        Map<String, String> result = new LinkedHashMap<>();
        result.put(APPROXIMATE_CLOSING_DATE_NCR_ONE,
                body.get(APPROXIMATE_CLOSING_DATE_VAL));
        result.put(NUMBER_OF_DAYS_LATE, body.get(NUMBER_OF_DAYS_LATE_VAL));
        result.put(NCR_DESCRIPTION, body.get(NCR_DESCRIPTION_VAL));
        result.put(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER,
                body.get(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER_VAL));
        result.put(RESPONSIBLE_PARTY_ROLE,
                body.get(RESPONSIBLE_PARTY_ROLE_VAL));
        result.put(IS_OPENED_BEFORE, body.get(IS_OPENED_BEFORE_VAL));
        result.put(OPENED_BEFORE_NCR_NO, body.get(OPENED_BEFORE_NCR_NO_VAL));
        result.put(CORRECTIVE_ACTION, body.get(CORRECTIVE_ACTION_VAL));
        result.put(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION,
                body.get(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL));
        result.put(TYPE_OF_TEST, body.get(TYPE_OF_TEST_VAL));
        result.put(REMARKS, body.get(REMARKS_VAL));
        return result;
    }

    private Map<String,String> getConfigTwoNCRBodyRows(
            Map<String, String> body) {
        Map<String, String> result = new LinkedHashMap<>();
        result.put(APPROXIMATE_CLOSING_DATE_NCR_TWO,
                body.get(APPROXIMATE_CLOSING_DATE_VAL));
        result.put(ORIGINAL_EXPECTED_CLOSING_DATE,
                body.get(ORIGINAL_EXPECTED_CLOSING_DATE_VAL));
        result.put(NUMBER_OF_DAYS_LATE,
                body.get(NUMBER_OF_DAYS_LATE_VAL));
        result.put(DAMAGE,
                body.get(DAMAGE_VAL));
        result.put(QUALITY_IMPACT,
                body.get(QUALITY_IMPACT_VAL));
        result.put(NCR_DESCRIPTION,
                body.get(NCR_DESCRIPTION_VAL));
        result.put(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER,
                body.get(RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER_VAL));
        result.put(CORRECTIVE_ACTION,
                body.get(CORRECTIVE_ACTION_VAL));
        result.put(RESPONSIBLE_PERSON,
                body.get(RESPONSIBLE_PERSON_VAL));
        result.put(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION,
                body.get(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL));
        result.put(REMARKS,
                body.get(REMARKS_VAL));
        return result;
    }
}
