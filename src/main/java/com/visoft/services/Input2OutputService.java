package com.visoft.services;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

@Service
public class Input2OutputService {

    StreamingResponseBody getOutput(InputStream inputStream){
        return outputStream -> {
            int nRead;
            byte[] data = new byte[10240];
            while ((nRead = inputStream.read(data, 0, data.length)) != -1) {
                outputStream.write(data, 0, nRead);
            }
        };
    }

    public StreamingResponseBody writeWorkBook(XSSFWorkbook workbook) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        InputStream inputStream = new ByteArrayInputStream(out.toByteArray());
        return getOutput(inputStream);
    }
}
