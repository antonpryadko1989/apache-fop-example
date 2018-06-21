package com.visoft.services;

import com.itextpdf.text.DocumentException;
import com.visoft.exceptions.FormatException;
import com.visoft.exceptions.TemplateValidationException;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.regulars.TemplateValue;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.mime.MediaType;
import org.apache.tika.parser.AutoDetectParser;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;
import java.util.Iterator;

@Service
public class POIXls {

    private static final String NOT_SUPPORTED_FORMAT = "Not supported format";

    @Value("${templates.repository}")
    private String templatesRepository;

    @Value("${app.XLSX}")
    private String appXLSX;

    @Value("${app.XLS}")
    private String appXLS;

    StreamingResponseBody getXLSXFile(TemplateDTO template) throws IOException, DocumentException {
        byte[] bytes = Files.readAllBytes(Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()));
        final ByteArrayInputStream is = new ByteArrayInputStream(bytes);
        final MediaType mediaType = new AutoDetectParser().getDetector().detect(is, new Metadata());
        if (mediaType.toString().equals(appXLS)) {
            return getXLS(template);
//        } else if (mediaType.toString().equals(appXLSX)) {
//            return getXLSX(template);
        } else
            throw new FormatException(NOT_SUPPORTED_FORMAT);    }


//        private static String getXLSX(TemplateDTO template) throws IOException {
//        String filePath = PATH_TO_TEMPLATE_REPOSITORY + File.separator + template.getProjectId() + File.separator + template.getTemplateName();
//        InputStream inputFile = new FileInputStream(filePath);
//        XSSFWorkbook wb = new XSSFWorkbook(inputFile);
//        XSSFSheet sheet = wb.getSheetAt(0);
//            for (Row row : sheet) {
//                Iterator<Cell> cellIterator = row.cellIterator();
//                while (cellIterator.hasNext()) {
//                    Cell cell = cellIterator.next();
//                    if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
//                        String newCellValueKey = TemplateValue.isValue(cell.getStringCellValue());
//                        if (newCellValueKey.length() > 0) {
//                            System.out.println(template.getTemplateBody().get(newCellValueKey));
//                            cell.setCellValue(template.getTemplateBody().get(newCellValueKey));
//                        }
//                    }
//                }
//            }
//        String fileName = "E:\\WORK\\TEMP\\" + new Date().getTime() + "_" + template.getOutPutName() ;
//        FileOutputStream out = new FileOutputStream(new File(fileName));
//        wb.write(out);
//        out.close();
//        return fileName;
//    }


    private StreamingResponseBody getXLS(TemplateDTO template) throws IOException {
        String filePath = Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()).toString();
        InputStream inputFile = new FileInputStream(filePath);
        HSSFWorkbook wb = new HSSFWorkbook(inputFile);
        HSSFSheet sheet = wb.getSheetAt(0);

        for (Row row : sheet) {
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    String newCellValueKey = TemplateValue.isValue(cell.getStringCellValue());
                    if (newCellValueKey.length() > 0) {
                        cell.setCellValue(template.getTemplateBody().get(newCellValueKey));
                    }
                }
            }
        }
        FileOutputStream out = new FileOutputStream(new File(fileName));
        wb.write(out);
        wb.write(out);
        out.close();
        return fileName;
    }

}
