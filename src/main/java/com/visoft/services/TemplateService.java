package com.visoft.services;

import com.visoft.exceptions.PreviewLogoException;
import com.visoft.exceptions.TemplateNameException;
import com.visoft.exceptions.TemplateValidationException;
import com.visoft.services.impl.*;
import com.visoft.templates.entity.TemplateDTO;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.annotation.Resource;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import static com.visoft.utils.Const.*;

@Service(value = "templateService")
public class TemplateService {

    private static final String NOT_EXIST = "Does not exist";
    private static final String NOT_ENOUGH_INFORMATION = "Not enough information ";
    private static final String TEMPLATE_NAME = TEMPLATE_NAME_VAL;


    @Resource
    private FOPPdf fopPdf;


    @Resource
    private POIXls poiXls;
    private static final String NOT_SUPPORTED = "not supported";

    public StreamingResponseBody getPDFFromTemplate(TemplateDTO template) {
        Map<String, String> templateValid = templateValidator(template);
        if(templateValid.isEmpty()) {
            return fopPdf.getPDFFile(template);
        }
        throw new TemplateValidationException(NOT_ENOUGH_INFORMATION, templateValid);
    }

//    public StreamingResponseBody getExcelFromTemplate(TemplateDTO template) {
//        Map<String, String> templateValid = templateValidator(template);
//        if(templateValid.isEmpty()) {
//            try {
//                return poiXls.getExcelFile(template);
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
//        throw new TemplateValidationException(NOT_ENOUGH_INFORMATION, templateValid);
//    }

    private Map<String, String> templateValidator(TemplateDTO template) {
        Map<String, String> templateValid = new HashMap<>();
        if (template.getTemplateName() == null || template.getTemplateName().equals("")) {
            templateValid.put("templateName", NOT_EXIST);
        }
        if (template.getConfigType() == null){
            templateValid.put("configType", NOT_EXIST);
        }
//        else if(!template.getConfigType().equals(ConfigType.ZERO) ||
//            !template.getConfigType().equals(ConfigType.ONE) || !template.getConfigType().equals(ConfigType.TWO)){
//            templateValid.put("configType", NOT_SUPPORTED);
//        }
        if (template.getProjectId() == null || template.getProjectId().equals("")) {
            templateValid.put("projectId", NOT_EXIST);
        }
        if (template.getOutPutName() == null || template.getOutPutName().equals("")) {
            templateValid.put("outPutName", NOT_EXIST);
        }
        if (template.getBody() == null) {
            templateValid.put("body", NOT_EXIST);
        }
        if(template.getSheetName() == null) {
            templateValid.put("sheetName", NOT_EXIST);
        }
        return templateValid;
    }

    public StreamingResponseBody previewXLSXLogo(int countCells,
                                                 MultipartFile multipartFile) {
        if (countCells < 3 || countCells > 8) {
            throw new PreviewLogoException("must be between 3-8", countCells);
        }
        byte[] outputByteArray = null;
        if (!multipartFile.isEmpty()) {
            try {
                return poiXls.addLogoToSheet(multipartFile.getBytes(),
                        countCells);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        throw new PreviewLogoException(NOT_EXIST);
    }

    public StreamingResponseBody downloadXLSX(TemplateDTO template) throws IOException {
        Map<String, String> templateValid = templateValidator(template);
        if(templateValid.isEmpty()) {
            XLSXBuilder builder;
            switch (template.getTemplateName().toUpperCase()){
                case NCR:
                    builder = new NCRReportImpl();
                    return builder.buildXLSX(template);
                case POC:
                    builder = new POCReportImpl();
                    return builder.buildXLSX(template);
                case APPROVAL_OF_SUBCONTRACTORS:
                    builder = new ApprovalOfSubcontractorsReportImpl();
                    return builder.buildXLSX(template);
                case APPROVAL_OF_SUPPLIERS:
                    builder = new ApprovalOfSuppliersReportImpl();
                    return builder.buildXLSX(template);
                case PRELIMINARY_MATERIALS_INSPECTION:
                    builder = new PreliminaryMaterialsInspectionReportImpl();
                    return builder.buildXLSX(template);
                case CHECKLIST:
                    builder = new ChecklistReportImpl();
                    return builder.buildXLSX(template);
                default:
                    throw new TemplateNameException(
                            template.getTemplateName(), NOT_SUPPORTED);
            }
        }
        throw new TemplateValidationException(NOT_ENOUGH_INFORMATION, templateValid);
    }
}