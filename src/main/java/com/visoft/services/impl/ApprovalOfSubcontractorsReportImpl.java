package com.visoft.services.impl;

import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;

import static com.visoft.services.PoiXSSFCellStyleService.getApprovalOfSubcontractorsReportStyles;
import static com.visoft.services.impl.ReportBodyConverter.convertReportBodyToReportBodyRows;
import static com.visoft.utils.Const.*;
import static com.visoft.services.POIXls.*;

public class ApprovalOfSubcontractorsReportImpl implements XLSXBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(TemplateDTO template) throws IOException {
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());
        XSSFCellStyle[] styles = getApprovalOfSubcontractorsReportStyles(xlsxObject);

        addLogo(
                xlsxObject, template.getBody()
        );
        addHeaderInfo(
                xlsxObject, APP_OF_SUBCONTR_TEMPLATE_NAME, VERSION, DATE,
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
        setValuesToRow(
                xlsxObject,
                convertReportBodyToReportBodyRows(
                        template.getBody().getBodyElements(),
                        APPROVAL_OF_SUBCONTRACTORS_BODY_KEYS,
                        APPROVAL_OF_SUBCONTRACTORS_BODY_VALUES
                ),
                styles[1],  styles[14], styles[11],
                styles[16], styles[12], styles[18]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(
                xlsxObject, CERTIFICATIONS,
                styles[0]
        );
        addCertifications(
                xlsxObject,
                CERTIFICATIONS_ITEM, CERTIFICATIONS_EXISTS, CERTIFICATIONS_CERTIFICATE_NO,
                CERTIFICATIONS_EXPIRATION, CERTIFICATIONS_ATTACHED_DOCUMENTS,
                styles[7], styles[8], styles[9]
        );
        if(notEmpty){
            addCertificationsList(
                    xlsxObject, template.getBody().getBodyLists().get(CERTIFICATIONS_VAL),
                    styles[1],  styles[13], styles[14], styles[11], styles[15], styles[16],
                    styles[12], styles[17], styles[18], styles[7],  styles[19], styles[20]
            );
        } else {
            addCertificationsList(
                    xlsxObject, null,
                    styles[1],  styles[13], styles[14], styles[11], styles[15], styles[16],
                    styles[12], styles[17], styles[18], styles[7],  styles[19], styles[20]
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

