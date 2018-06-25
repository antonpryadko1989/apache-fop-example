package com.visoft.api;

import com.visoft.services.TemplateService;
import com.visoft.templates.entity.TemplateDTO;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.annotation.Resource;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.file.Paths;

@Controller
@RequestMapping(value = "/visoft/api")
public class TemplateController {

    @Value("${templatesPath}")
    private String templatesPath;

    @Value("${app.xls}")
    private String appXLS;

    @Value("${app.xlsx}")
    private String appXLSX;

    @Resource
    private TemplateService templateService;

    @RequestMapping(value = "/downloadPdf", method = RequestMethod.POST,
            headers = {"Accept=application/json"})
    public StreamingResponseBody downloadPdf(HttpServletResponse response, @RequestBody TemplateDTO template) {
        response.setContentType("application/pdf; filename=" + template.getOutPutName());
        response.setHeader("filename", template.getOutPutName());
        return templateService.getPDFFromTemplate(template);
            }

    @RequestMapping(value = "/downloadExcel", method = RequestMethod.POST, headers = {"Accept=application/json"})
    public StreamingResponseBody getExcelFromTemplate(final HttpServletResponse response, @RequestBody TemplateDTO template) {
        response.setHeader("Content-Disposition", "attachment; filename=" + template.getOutPutName());
        response.setContentType("application/octet-stream; filename=" + template.getOutPutName() + "; charset=UTF-8");
        response.setHeader("filename", template.getOutPutName());
        return templateService.getExcelFromTemplate(template);
    }

    @RequestMapping(value = "/download/{projectId}/{path}", method = RequestMethod.GET)
    public void download(HttpServletResponse response, @PathVariable("projectId") String projectId, @PathVariable("path") String path) throws IOException {
        FileInputStream fileIn = new FileInputStream(Paths.get(templatesPath, projectId, path).toFile());
        ServletOutputStream out = response.getOutputStream();

        byte[] outputByte = new byte[4096];
        while(fileIn.read(outputByte, 0, 4096) != -1){
            out.write(outputByte, 0, 4096);
        }
        fileIn.close();
        out.close();
    }
}
