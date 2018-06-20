package com.visoft.api;

import com.visoft.exceptions.PathValidationException;
import com.visoft.services.TemplateService;
import com.visoft.templates.entity.TemplateDTO;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;

@RestController
@RequestMapping(value = "/visoft/api")
public class TemplateController {

    @Value("${templatesPath}")
    private String templatesPath;

    @Resource
    private TemplateService templateService;

    // CREATE PDF FORM OBJECT & RETURN PATH TO PDF
//    @RequestMapping(value = "/getURL2Pdf", method = RequestMethod.POST, headers = {"Accept=application/json"}, produces = {"application/json; charset=UTF-8"})
//    public ResponseEntity<String > getPdf(@RequestBody TemplateDTO template) {
//        return templateService.getPDFFromTemplate(template);
//    }


//    @RequestMapping(value = "/previewFile", method = RequestMethod.GET)
//    public void getFile(
//            @RequestParam(value ="fileName") String fileName,
//            HttpServletResponse response) {
//        try {
//            InputStream is = new FileInputStream(fileName);
//            org.apache.commons.io.IOUtils.copy(is, response.getOutputStream());
//            response.flushBuffer();
//        } catch (IOException ex) {
//            throw new RuntimeException("IOError writing file to output stream");
//        }
//    }

//    @RequestMapping(value = "/download/{templateName}", method = RequestMethod.GET)
//    public void download(final HttpServletResponse response, @PathVariable("templateName") String templateName) throws IOException {
//
//        File file = new File(templatesPath + templateName + ".XLSX");
//
//        response.setContentType("application/vnd.ms-excel");
//        response.setHeader("Content-Disposition", "attachment; filename=" + file.getName());
//        response.getOutputStream().flush();
//    }

//    @RequestMapping(value = "downloadFile", method = RequestMethod.GET)
//    public StreamingResponseBody getSteamingFile(HttpServletResponse response,
//                                                 @RequestParam(value = "path") String path) throws IOException {
//        response.setContentType("application/pdf");
//        int iend = path.indexOf("_");
//        String subString = path.substring(iend + 1, path.length());
//        response.setHeader("Content-Disposition", "attachment; filename=\"" + subString + "\"");
//        try {
//            return templateService.downloadPDF(path);
//
//        }catch (IOException e){
//            throw new PathValidationException("SOMETHING ERRORS");
//
//        }
//        finally {
//            Files.delete(Paths.get(path));
//        }
//    }

//
//    @RequestMapping(value = "/getPdf", method = RequestMethod.GET, produces = "application/pdf")
//    public
//    ResponseEntity<InputStreamResource> downloadPdfFile
//            (@RequestParam(value = "path") String path)
//            throws IOException {
//        File pdfFile = Paths.get(path).toFile();
//
//        return ResponseEntity.ok().contentLength(pdfFile.length())
//                .contentType(MediaType.parseMediaType("application/octer-stream"))
//                .body(new InputStreamResource(new FileInputStream(pdfFile), ".pdf"));
//    }


//    @RequestMapping(value = "/pdf", method = RequestMethod.GET, produces = "application/pdf")
//    public ResponseEntity<InputStreamResource> download(@RequestParam(value = "fileName") String fileName) throws IOException {
//        System.out.println("Calling Download:- " + fileName);
//        File pdfFile = Paths.get(fileName).toFile();
//        HttpHeaders headers = new HttpHeaders();
//        headers.setContentType(MediaType.parseMediaType("application/pdf"));
//        headers.add("Access-Control-Allow-Origin", "*");
//        headers.add("Access-Control-Allow-Methods", "GET, POST, PUT");
//        headers.add("Access-Control-Allow-Headers", "Content-Type");
//        headers.add("Content-Disposition", "filename=" + fileName);
//        headers.add("Cache-Control", "no-cache, no-store, must-revalidate");
//        headers.add("Pragma", "no-cache");
//        headers.add("Expires", "0");
//
//        headers.setContentLength(pdfFile.length());
//        ResponseEntity<InputStreamResource> response = new ResponseEntity<>(
//                new InputStreamResource(new FileInputStream(fileName)), headers, HttpStatus.OK);
//        return response;
//
//    }
    @RequestMapping(value = "/downloadPdf", method = RequestMethod.POST,
            headers = {"Accept=application/json"}, produces = {"application/pdf; charset=UTF-8"})
    public StreamingResponseBody downloadPdf(HttpServletResponse response,
                                             @RequestBody TemplateDTO template) {
        response.setContentType("application/pdf");
                return templateService.getPDFFromTemplate(template);
}


}
