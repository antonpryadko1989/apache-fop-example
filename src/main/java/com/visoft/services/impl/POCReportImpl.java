package com.visoft.services.impl;

import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;

import static com.visoft.services.POIXls.*;
import static com.visoft.services.PoiXSSFCellStyleService.getPOCReportStyles;
import static com.visoft.services.impl.ReportBodyConverter.convertReportBodyToReportBodyRows;
import static com.visoft.utils.Const.*;

public class POCReportImpl implements XLSXBuilder{

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(TemplateDTO template) throws IOException {
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());
        XSSFCellStyle styles[] = getPOCReportStyles(xlsxObject);

        addLogo(
                xlsxObject, template.getBody()
        );

        addHeaderInfo(
                xlsxObject, POC_TEMPLATE_NAME, VERSION, DATE,
                template.getBody().getBodyElements().get(TEMPLATE_NAME_VAL),
                template.getBody().getBodyElements().get(VERSION_VAL),
                template.getBody().getBodyElements().get(DATE_VAL),
                styles[1], styles[2], styles[3], styles[12], styles[4], styles[5]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(
                xlsxObject, template,
                styles[1],  styles[13], styles[2],  styles[14], styles[11], styles[15],
                styles[10], styles[16], styles[6],  styles[17], styles[4],  styles[18]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        if(notEmpty) {
            if (template.getBody().getBodyLists().get(DRAWINGS_VAL) != null) {
                addDrawings(
                        xlsxObject, DRAWING_NO, DRAWING_VERSION_REVISION, DRAWING_NAME,
                        styles[7], styles[8], styles[9]
                );
                addDrawingsList(
                        xlsxObject, template.getBody().getBodyLists().get(DRAWINGS_VAL),
                        styles[19], styles[13], styles[14], styles[20], styles[15], styles[16],
                        styles[21], styles[17], styles[18], styles[22], styles[23], styles[24]
                );
                addEmptyStringWithGivenHeight(xlsxObject, 9);
            }
        }
        setValuesToRow(
                xlsxObject,
                convertReportBodyToReportBodyRows(
                        template.getBody().getBodyElements(),
                        POC_BODY_KEYS,
                        POC_BODY_VALUES
                ),
                styles[1],  styles[14], styles[11],
                styles[16], styles[12], styles[18]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(
                xlsxObject, ADDITIONAL_DOCUMENTS,
                styles[0]
        );
        addAdditionalDocuments(
                xlsxObject, ADDITIONAL_DOCUMENTS_ITEM, ADDITIONAL_DOCUMENTS_EXISTS, ADDITIONAL_DOCUMENTS_CERTIFICATE_NO,
                ADDITIONAL_DOCUMENTS_EXPIRATION, ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS,
                styles[7], styles[8], styles[9]
        );
        if(notEmpty){
            addAdditionalDocumentsList(
                    xlsxObject, template.getBody().getBodyLists().get(ADDITIONAL_DOCUMENTS_VAL),
                    styles[1],  styles[13], styles[14], styles[11], styles[15], styles[16],
                    styles[12], styles[17], styles[18], styles[7],  styles[23], styles[24]
            );
        }else{
            addAdditionalDocumentsList(
                    xlsxObject, null,
                    styles[1],  styles[13], styles[14], styles[11], styles[15], styles[16],
                    styles[12], styles[17], styles[18], styles[7],  styles[23], styles[24]
            );
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(
                xlsxObject, APPROVALS,
                styles[0]
        );
        addApprovals(
                xlsxObject, APPROVALS_ROLE, APPROVALS_NAME, APPROVALS_SIGNATURE, APPROVALS_DATE, APPROVALS_STATUS,
                styles[7], styles[8], styles[9]
        );
        if(notEmpty){
            addApprovalsList(
                    xlsxObject, template.getBody().getBodyLists().get(APPROVALS_VAL),
                    styles[1],  styles[13], styles[14], styles[11], styles[15],
                    styles[16], styles[12], styles[17], styles[18]
            );
        }else {
            addApprovalsList(
                    xlsxObject, null,
                    styles[1],  styles[13], styles[14], styles[11], styles[15],
                    styles[16], styles[12], styles[17], styles[18]
            );
        }
        return input2OutputService.writeWorkBook(workbook);
    }
}