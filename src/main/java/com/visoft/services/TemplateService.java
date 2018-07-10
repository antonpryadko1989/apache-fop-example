package com.visoft.services;

import com.visoft.exceptions.TemplateValidationException;
import com.visoft.templates.entity.TemplateDTO;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;


import javax.annotation.Resource;
import java.io.IOException;
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
        if(templateValid.isEmpty()) {
            return fopPdf.getPDFFile(template);
        }
        throw new TemplateValidationException(NOT_ENOUGH_INFORMATION, templateValid);
    }

    public StreamingResponseBody getExcelFromTemplate(TemplateDTO template) {
        Map<String, String> templateValid = templateValidator(template);
        if(templateValid.isEmpty()) {
            try {
                return poiXls.getExcelFile(template);
            } catch (IOException e) {
                e.printStackTrace();
            }
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
        if (template.getBody() == null) {
            templateValid.put("body", NOT_EXIST);
        }
        return templateValid;
    }
}