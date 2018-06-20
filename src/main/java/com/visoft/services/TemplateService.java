package com.visoft.services;

import com.visoft.exceptions.TemplateValidationException;
import com.visoft.templates.entity.TemplateDTO;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.annotation.Resource;
import java.io.*;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

@Service(value = "templateService")
public class TemplateService {

    @Resource
    private FOPPdfDemo fopPdfDemo;

    private HttpHeaders responseHeaders = new HttpHeaders();

    public StreamingResponseBody getPDFFromTemplate(TemplateDTO template) {
        Map<String, String> templateValid = templateValidator(template);
        if(templateValid.isEmpty()) {
            return fopPdfDemo.getPDFUrl(template);
        }
        throw new TemplateValidationException("I need more information ", templateValid);
    }
//
//    public StreamingResponseBody downloadPDF(String path) throws IOException {
//        InputStream inputStream = new FileInputStream(Paths.get(path).toFile());
//        return outputStream -> {
//            int nRead;
//            byte[] data = new byte[1024];
//            while ((nRead = inputStream.read(data, 0, data.length)) != -1) {
//                outputStream.write(data, 0, nRead);
//            }
//        };
//    }

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
