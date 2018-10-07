package com.visoft;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import static com.visoft.cellStyleUtil.CellStyleUtil.setAllBordersByBorderStyle;
import static com.visoft.services.POIXls.setILColumnWidths;

@SpringBootApplication
public class Application {

    public static void main(String[] args) {
        SpringApplication.run(Application.class, args);
    }
}