package com.visoft.api;

import com.visoft.services.TemplateService;
import com.visoft.templates.entity.TemplateDTO;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
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

    @RequestMapping(value = "/downloadXLSX",
            method = RequestMethod.POST)
    public StreamingResponseBody downloadXLSX (
            HttpServletResponse response,
            @RequestBody TemplateDTO template) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;");
        response.setHeader("filename", template.getOutPutName());
        return templateService.downloadXLSX(template);
    }

    @RequestMapping(value = "/previewXLSXLogo",
            method = RequestMethod.POST)
    public StreamingResponseBody previewXLSXLogo (
            HttpServletResponse response,
            @RequestParam(value = "countCells") int countCells,
            @RequestPart(value = "logo") MultipartFile multipartFile) {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;");
        response.setHeader("filename", "previewXLSXLogo");
        return templateService.previewXLSXLogo(countCells, multipartFile);
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
