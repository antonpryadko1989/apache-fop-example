package com.visoft.services.impl;

import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import com.visoft.templates.enums.ConfigType;
import org.apache.commons.collections4.map.LinkedMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;
import java.util.Map;

import static com.visoft.services.POIXls.*;
import static com.visoft.services.PoiXSSFCellStyleService.getNCRReportStyles;
import static com.visoft.services.impl.ReportBodyConverter.convertReportBodyToReportBodyRows;
import static com.visoft.utils.Const.*;

public class NCRReportImpl implements XLSXBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(TemplateDTO template) throws IOException {
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());
        XSSFCellStyle styles[] = getNCRReportStyles(xlsxObject);
        addLogo(
                xlsxObject, template.getBody()
        );
        addHeaderInfo(
                xlsxObject, NCR_TEMPLATE_NAME, VERSION, DATE,
                template.getBody().getBodyElements().get(REPORT_NAME),
                template.getBody().getBodyElements().get(VERSION_VAL),
                template.getBody().getBodyElements().get(DATE_VAL),
                styles[2],  styles[3], styles[4],
                styles[13], styles[5], styles[6]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(
                xlsxObject, template,
                styles[2],  styles[14], styles[3],  styles[15], styles[12], styles[16],
                styles[11], styles[17], styles[7],  styles[18], styles[5],  styles[19]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        setNCRNumber(
                xlsxObject, NCR_NUMBER, template.getBody().getBodyElements().get(NCR_NUMBER_VAL),
                styles[8], styles[22]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addQaQc(
                xlsxObject, NCR_OPENED, DATE_OF_NCR,
                template.getBody().getBodyElements().get(QC_OPENED_NAME_VAL),
                template.getBody().getBodyElements().get(QC_OPENED_DATE_OF_NCR_VAL),
                template.getBody().getBodyElements().get(QA_OPENED_NAME_VAL),
                template.getBody().getBodyElements().get(QA_OPENED_DATE_OF_NCR_VAL),
                styles[8],  styles[9], styles[10], styles[3], styles[14],
                styles[15], styles[5], styles[18], styles[19]
        );
        addAlignmentRow(
                xlsxObject, DETAILS_OF_STRUCTURE, styles[0]
        );
        addNCRBody(
                xlsxObject, template,
                styles[2],  styles[3],  styles[4],  styles[20], styles[16],
                styles[17], styles[12], styles[13], styles[19]
        );
        addAlignmentRow(
                xlsxObject, ACCEPTANCE_OF_CORRECTIVE_ACTION, styles[0]
        );
        addQaQc(
                xlsxObject, NCR_CLOSED, CLOSING_DATE,
                template.getBody().getBodyElements().get(QC_CLOSED_NAME_VAL),
                template.getBody().getBodyElements().get(QC_CLOSED_DATE_OF_NCR_VAL),
                template.getBody().getBodyElements().get(QA_CLOSED_NAME_VAL),
                template.getBody().getBodyElements().get(QA_CLOSED_DATE_OF_NCR_VAL),
                styles[8], styles[9], styles[10], styles[3],styles[14],
                styles[15],styles[5], styles[18], styles[19]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(
                xlsxObject, ADDITIONAL_DOCUMENTS, styles[1]
        );
        addAdditionalDocuments(
                xlsxObject, ADDITIONAL_DOCUMENTS_ITEM, ADDITIONAL_DOCUMENTS_EXISTS,
                ADDITIONAL_DOCUMENTS_CERTIFICATE_NO, ADDITIONAL_DOCUMENTS_EXPIRATION, ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS,
                styles[8], styles[9], styles[10]
        );
        if(notEmpty) {
            addAdditionalDocumentsList(
                    xlsxObject, template.getBody().getBodyLists().get(ADDITIONAL_DOCUMENTS_VAL),
                    styles[2],  styles[14], styles[15], styles[12], styles[17], styles[17],
                    styles[13], styles[18], styles[19], styles[8],  styles[21], styles[22]
            );
        } else {
            addAdditionalDocumentsList(
                    xlsxObject, null,
                    styles[2],  styles[14], styles[15], styles[12], styles[17], styles[17],
                    styles[13], styles[18], styles[19], styles[8],  styles[21], styles[22]
            );
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
                return convertReportBodyToReportBodyRows(body, CONFIG_ZERO_NCR_BODY_KEYS, CONFIG_ZERO_NCR_BODY_VALUES);
            case ONE:
                return convertReportBodyToReportBodyRows(body, CONFIG_ONE_NCR_BODY_KEYS, CONFIG_ONE_NCR_BODY_VALUES);
            default:
                return convertReportBodyToReportBodyRows(body, CONFIG_TWO_NCR_BODY_KEYS, CONFIG_TWO_NCR_BODY_VALUES);
        }
    }
}