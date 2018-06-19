package com.visoft.services;

import com.visoft.exceptions.TemplateValidationException;
import com.visoft.templates.entity.TemplateDTO;
import org.json.JSONObject;
import org.json.XML;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

@Service(value = "templateService")
public class TemplateService {

    @Resource
    private FOPPdfDemo fopPdfDemo;

    private HttpHeaders responseHeaders = new HttpHeaders();

    public ResponseEntity<String> getPDFFromTemplate(TemplateDTO template) {
        Map<String, String> templateValid = templateValidator(template);
        if(templateValid.isEmpty()) {
            return new ResponseEntity<>(fopPdfDemo.getPDFUrl(template), HttpStatus.OK);
        }
        throw new TemplateValidationException("I need more information ", templateValid);
    }

    private HashMap<String, String> templateValidator(TemplateDTO template) {
        HashMap<String, String> userValid = new HashMap<>();
        if (template.getTemplateName() == null || template.getTemplateName().equals("")) {
            userValid.put("templateName", "Does not exist");
        }
        if (template.getProjectId() == null || template.getProjectId().equals("")) {
            userValid.put("projectId", "Does not exist");
        }
        if (template.getOutPutName() == null || template.getOutPutName().equals("")) {
            userValid.put("outPutName", "Does not exist");
        }
        if (template.getBody() == null || template.getBody().equals("")) {
            userValid.put("body", "Does not exist");
        }
        return userValid;
    }
}
