package com.visoft.api;

import com.itextpdf.text.DocumentException;
import com.visoft.services.TemplateService;
import com.visoft.templates.entity.TemplateDTO;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.MediaType;
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
            headers = {"Accept=application/json"}, produces = {"application/pdf; charset=UTF-8"})
    public StreamingResponseBody downloadPdf(HttpServletResponse response, @RequestBody TemplateDTO template) {
        response.setContentType("application/pdf; filename=" + template.getOutPutName());
        response.setHeader("filename", template.getOutPutName());
        return templateService.getPDFFromTemplate(template);
            }

    @RequestMapping(value = "/downloadExcel", method = RequestMethod.POST,
            headers = {"Accept=application/json"}, produces = {"application/vnd.ms-excel; charset=UTF-8"})
    public StreamingResponseBody downloadXLSX(final HttpServletResponse response, @RequestBody TemplateDTO template) throws IOException, DocumentException {
        response.setHeader("Content-Disposition", "attachment; filename=" + template.getOutPutName());
        response.setContentType("application/xml; filename=" + template.getOutPutName() + "; charset=UTF-8");
        response.setHeader("filename", template.getOutPutName());
        return templateService.getExcelFromTemplate(template);
    }

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void download(final HttpServletResponse response, @RequestParam("templateName") String templateName) throws IOException {
        response.setHeader("Content-Disposition", "attachment; filename=" + templateName);
        FileInputStream fileIn = new FileInputStream(Paths.get("E:/WORK/TEMP/NCR.xlsx").toFile());
        ServletOutputStream out = response.getOutputStream();

        byte[] outputByte = new byte[4096];
        while(fileIn.read(outputByte, 0, 4096) != -1){
            out.write(outputByte, 0, 4096);
        }
        fileIn.close();
        out.close();

    }


}
