package com.visoft.api;

import com.visoft.services.TemplateService;
import com.visoft.templates.entity.TemplateDTO;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import javax.annotation.Resource;

@RestController
@RequestMapping(value = "/visoft/api")
public class TemplateController {

    @Resource
    private TemplateService templateService;

    @RequestMapping(value = "/getURL2Pdf", method = RequestMethod.POST, headers = {"Accept=application/json"}, produces = {"application/json; charset=UTF-8"})
    public ResponseEntity<String > getPdf(@RequestBody TemplateDTO template) {
        return templateService.getPDFFromTemplate(template);
    }
}
