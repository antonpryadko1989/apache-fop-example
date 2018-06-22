package com.visoft.services;

import com.itextpdf.text.DocumentException;
import com.visoft.exceptions.TemplateValidationException;
import com.visoft.templates.entity.TemplateDTO;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.annotation.Resource;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

@Service(value = "templateService")
public class TemplateService {

    private static final String NOT_EXIST = "Does not exist";
    private static final String NOT_ENOUGH_INFORMATION = "Not enough information ";

    @Resource
    private FOPPdf fopPdf;


    @Resource
    private POIXls poiXls;

    public StreamingResponseBody getPDFFromTemplate(TemplateDTO template) {
        Map<String, String> templateValid = templateValidator(template);
        templateValid.putAll(checkBody(template));
        if(templateValid.isEmpty()) {
            return fopPdf.getPDFFile(template);
        }
        throw new TemplateValidationException(NOT_ENOUGH_INFORMATION, templateValid);
    }

    public StreamingResponseBody getExcelFromTemplate(TemplateDTO template) throws IOException, DocumentException {
        Map<String, String> templateValid = templateValidator(template);
        templateValid.putAll(checkTemplateBody(template));
        if(templateValid.isEmpty()) {
            return poiXls.getExcelFile(template);
        }
        throw new TemplateValidationException(NOT_ENOUGH_INFORMATION, templateValid);
    }

    private Map<String, String> templateValidator(TemplateDTO template) {
        Map<String, String> templateValid = new HashMap<>();
        if (template.getTemplateName() == null || template.getTemplateName().equals("")) {
            templateValid.put("templateName", NOT_EXIST);
        }
        if (template.getProjectId() == null || template.getProjectId().equals("")) {
            templateValid.put("projectId", NOT_EXIST);
        }
        if (template.getOutPutName() == null || template.getOutPutName().equals("")) {
            templateValid.put("outPutName", NOT_EXIST);
        }
        return templateValid;
    }

    private Map<String, String> checkBody(TemplateDTO template) {
        Map<String, String> bodyValid = new HashMap<>();
        if (template.getBody() == null || template.getBody().equals("")) {
            bodyValid.put("body", NOT_EXIST);
        }
        return bodyValid;
    }

    private Map<String, String> checkTemplateBody(TemplateDTO template) {
        Map<String, String> bodyValid = new HashMap<>();
        if (template.getTemplateBody()== null || template.getTemplateBody().isEmpty()) {
            bodyValid.put("templateBody", NOT_EXIST);
        }
        return bodyValid;
    }

}
