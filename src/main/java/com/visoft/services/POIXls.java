package com.visoft.services;

import com.visoft.exceptions.PathValidationException;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.regulars.TemplateValue;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.mime.MediaType;
import org.apache.tika.parser.AutoDetectParser;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.*;
import java.nio.file.*;
import java.util.*;

@Service
public class POIXls {

    @Value("${templates.repository}")
    private String templatesRepository;

    @Value("${app.XLSX}")
    private String appXLSX;

    @Value("${app.XLS}")
    private String appXLS;

    StreamingResponseBody getExcelFile(TemplateDTO template) throws IOException {
        File f = Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()).toFile();
        if (f.isDirectory()) {
            throw new PathValidationException("error", "It's directory");
        }
        if (!f.exists()) {
            throw new PathValidationException("error", "File not exist");
        }
        Workbook workbook = getWorkbook(Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()).toString());
        Sheet sheet = workbook.getSheetAt(0);
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
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        InputStream inputStream = new ByteArrayInputStream(out.toByteArray());
        return outputStream -> {
            int nRead;
            byte[] data = new byte[1024];
            while ((nRead = inputStream.read(data, 0, data.length)) != -1) {
                outputStream.write(data, 0, nRead);
            }
        };
    }

    private Workbook getWorkbook(String filePath)
            throws IOException {
        InputStream inputFile = new FileInputStream(filePath);
        byte[] bytes = Files.readAllBytes(Paths.get(filePath));
        final ByteArrayInputStream is = new ByteArrayInputStream(bytes);
        final MediaType mediaType = new AutoDetectParser().getDetector().detect(is, new Metadata());
        Workbook workbook = null;
        if (mediaType.toString().equals(appXLSX)) {
            workbook = new XSSFWorkbook(inputFile);
            return  workbook;
        } else if (mediaType.toString().equals(appXLS)) {
            workbook = new HSSFWorkbook(inputFile);
            return workbook;
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }
    }
}