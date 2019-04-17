package com.visoft.services.impl;

import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;

import static com.visoft.services.PoiXSSFCellStyleService.getPreliminaryMaterialsInspectionReportStyles;
import static com.visoft.services.impl.ReportBodyConverter.convertReportBodyToReportBodyRows;
import static com.visoft.utils.Const.*;
import static com.visoft.services.POIXls.*;

public class PreliminaryMaterialsInspectionReportImpl implements XLSXBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(TemplateDTO template) throws IOException {
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());
        XSSFCellStyle[] styles = getPreliminaryMaterialsInspectionReportStyles(xlsxObject);

        addLogo(
                xlsxObject, template.getBody()
        );
        addHeaderInfo(
                xlsxObject, PRELIM_MATERIALS_INS_TEMPLATE_NAME, VERSION, DATE, template.getBody().getBodyElements().get(TEMPLATE_NAME_VAL),
                template.getBody().getBodyElements().get(VERSION_VAL), template.getBody().getBodyElements().get(DATE_VAL),
                styles[1], styles[2], styles[3], styles[5], styles[4], styles[6]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(
                xlsxObject,template,
                styles[1],  styles[17], styles[2],  styles[18], styles[7],  styles[21],
                styles[8],  styles[20], styles[9],  styles[19], styles[4],  styles[16]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        setValuesToRow(
                xlsxObject,
                convertReportBodyToReportBodyRows(
                        template.getBody().getBodyElements(),
                        PRELIM_INSPEC_RESULT_BODY_KEYS,
                        PRELIM_INSPEC_RESULT_BODY_VALUES
                        ),
                styles[1], styles[18], styles[7], styles[20], styles[5], styles[16]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(
                xlsxObject, PRELIMINARY_INSPECTION_RESULTS,
                styles[0]
        );
        addPreliminaryInspectionResults(xlsxObject,
                PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION, PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS,
                PRELIM_INSPEC_RESULT_INSPECTION_RESULTS, PRELIM_INSPEC_RESULT_CERTIFICATE_NO,
                PRELIM_INSPEC_RESULT_PASS_FAIL, PRELIM_INSPEC_RESULT_COMMENTS,
                styles[10], styles[11], styles[12]
        );
        if (notEmpty) {
            addPreliminaryInspectionResultsList(
                    xlsxObject, template.getBody().getBodyLists().get(PRELIMINARY_INSPECTION_RESULTS_VAL),
                    styles[1], styles[17], styles[18], styles[7],  styles[21], styles[20],
                    styles[5], styles[19], styles[16], styles[10], styles[14], styles[15]
            );
        } else {
            addPreliminaryInspectionResultsList(
                    xlsxObject, null,
                    styles[1],  styles[17], styles[18], styles[7],  styles[21], styles[20],
                    styles[5],  styles[19], styles[16], styles[10], styles[14], styles[15]
            );
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(
                xlsxObject, ADDITIONAL_DOCUMENTS,
                styles[0]
        );

        addAdditionalDocuments(
                xlsxObject, ADDITIONAL_DOCUMENTS_ITEM,
                ADDITIONAL_DOCUMENTS_EXISTS, ADDITIONAL_DOCUMENTS_CERTIFICATE_NO,
                ADDITIONAL_DOCUMENTS_EXPIRATION, ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS,
                styles[10], styles[11], styles[13]
        );
        if (notEmpty) {
            addAdditionalDocumentsList(
                    xlsxObject, template.getBody().getBodyLists().get(ADDITIONAL_DOCUMENTS_VAL),
                    styles[1],  styles[17], styles[18], styles[7],  styles[21], styles[20],
                    styles[5],  styles[19], styles[16], styles[10], styles[14], styles[15]
            );
        } else {
            addAdditionalDocumentsList(
                    xlsxObject, null,
                    styles[1],  styles[17], styles[18], styles[7], styles[21], styles[20],
                    styles[5],  styles[19], styles[16], styles[10], styles[14], styles[15]
            );
        }
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addAlignmentRow(
                xlsxObject, APPROVALS,
                styles[0]
        );
        addApprovals(
                xlsxObject, APPROVALS_ROLE, APPROVALS_NAME, APPROVALS_SIGNATURE, APPROVALS_DATE, APPROVALS_STATUS,
                styles[10], styles[11], styles[13]
        );
        if (notEmpty) {
            addApprovalsList(
                    xlsxObject, template.getBody().getBodyLists().get(APPROVALS_VAL),
                    styles[1],  styles[17], styles[18], styles[7], styles[21],
                    styles[20], styles[5],  styles[19], styles[16]
            );
        } else {
            addApprovalsList(
                    xlsxObject, null,
                    styles[1],  styles[17], styles[18], styles[7], styles[21],
                    styles[20], styles[5],  styles[19], styles[16]
            );
        }
        return input2OutputService.writeWorkBook(workbook);
    }
}