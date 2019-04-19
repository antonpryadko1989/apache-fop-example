package com.visoft.services;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.RectangleReadOnly;
import com.itextpdf.text.pdf.PdfWriter;
import com.visoft.dto.DeficienciesReportDTO;
import com.visoft.exceptions.PathValidationException;
import net.sf.jmimemagic.MagicException;
import net.sf.jmimemagic.MagicMatchNotFoundException;
import net.sf.jmimemagic.MagicParseException;
import org.apache.fop.apps.*;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.json.JSONObject;
import org.json.XML;
import org.springframework.context.annotation.PropertySource;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.xml.transform.TransformerFactory;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamSource;
import java.io.*;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;

import static com.visoft.services.Extensions.pdfExtension;
import static com.visoft.utils.MineTypeValidator.getMineTypeFromByteArray;
import static java.nio.charset.StandardCharsets.UTF_8;
import static jdk.nashorn.internal.objects.NativeMath.log;


@PropertySource("classpath:application.properties")
public class PDFService {

    private static final String templatesRepository = "C:/Program Files/downloadDocsConfig";
    private static final String xmlStart = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><root>";
    private static final String xmlEnd = "</root>";
    private static final String fopExtension = ".xsl";
    private static final String config = "userconfig.xml";

    public StreamingResponseBody generatePDF(DeficienciesReportDTO reportDTO) {
        File xsltFile = getFile(reportDTO.getReportName() + fopExtension);
        File userConfig = getFile(config);
        if (xsltFile != null && userConfig != null) {
            try {
                byte[] reportPdf = convertPDF(reportDTO, xsltFile, userConfig);
                if(reportPdf.length>10){
                    if(reportDTO.getDocs()!=null&&!reportDTO.getDocs().isEmpty()){
                        return concatPDF(reportPdf, reportDTO.getDocs());
                    }else return  new Input2OutputService().getOutput(new ByteArrayInputStream(reportPdf));
                }else throw new PathValidationException("error", "Something get wrong");
            } catch (Exception e) {
                e.printStackTrace();
                throw new PathValidationException("error", "Something get wrong");
            }
        } else throw new PathValidationException("error", "Something get wrong");
    }


    private File getFile(String name) {
        File file = Paths.get(templatesRepository, name).toFile();
        return !file.isDirectory() && file.exists() ? file : null;
    }

    private byte[] convertPDF(DeficienciesReportDTO reportDTO,
                              File xsltFile,
                              File userConfig) throws Exception {
        FopFactory fopFactory = FopFactory.newInstance(userConfig);
        StreamSource xmlSource = new StreamSource(
                new ByteArrayInputStream(
                        (xmlStart + XML.toString(
                                new JSONObject(reportDTO)
                        ) + xmlEnd)
                                .getBytes(UTF_8)
                )
        );
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()){
            Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, fopFactory.newFOUserAgent(), out);
            TransformerFactory.newInstance()
                    .newTransformer(new StreamSource(xsltFile))
                    .transform(xmlSource,new SAXResult(fop.getDefaultHandler()));
            out.flush();
            return out.toByteArray();
        }catch (Exception e){
            throw new Exception();
        }
    }

    private StreamingResponseBody concatPDF(byte[] reportPdf, List<byte[]> docs) throws IOException {
        List<byte[]> pdfList = new ArrayList<>();
        pdfList.add(reportPdf);
        for (byte[] b:docs) {
            if(b!=null){
                byte[] isPdf = checkMineType(b);
                if(isPdf!=null){
                    pdfList.add(isPdf);
                }
            }
        }
        if(pdfList.size()>1){
            return new Input2OutputService().getOutput(new ByteArrayInputStream(concatPdf(pdfList).toByteArray()));
        }else{
            return new Input2OutputService().getOutput(new ByteArrayInputStream(pdfList.get(0)));
        }
    }

    private byte[] checkMineType(byte[] b) {
        try {
            if(checkPDFMineType(b)){
                return b;
            } else if(checkImageMineType(b)){
                return convertImageToPDF(b);
            }
        }catch (Throwable e){
            return null;
        }
        return null;
    }

    private byte[] convertImageToPDF(byte[] b) {
        byte[] result = null;
        try{
            Document document;
            ByteArrayOutputStream stream = new ByteArrayOutputStream();
            Image img = Image.getInstance(b);
            float imgHeight = img.getHeight();
            float imgWidth = img.getWidth();
            if(imgWidth > imgHeight){
                document = new Document(new RectangleReadOnly(imgWidth, imgHeight, 90), 0,0,0,0);
            }else {
                document = new Document(new RectangleReadOnly(imgWidth, imgHeight,0), 0,0,0,0);
            }
            PdfWriter.getInstance(document, stream);
            document.open();
            document.add(img);
            document.close();
            result = stream.toByteArray();
        } catch (IOException | DocumentException e) {
            log(Level.WARNING, e.getMessage());
        }
        return result;
    }

    private boolean checkImageMineType(byte[] b) {
        boolean result = false;
        try {
            String ext = getMineTypeFromByteArray(b);
            result = ext.contains(getMineTypeFromByteArray(b));
        } catch (MagicException | MagicParseException | MagicMatchNotFoundException e) {
            e.printStackTrace();
        }
        return result;
    }

    private boolean checkPDFMineType(byte[] b){
        boolean result = false;
        try {
            String ext = getMineTypeFromByteArray(b);
            result = ext.equals(pdfExtension);
        } catch (MagicException | MagicParseException | MagicMatchNotFoundException e) {
            e.printStackTrace();
        }
        return result;
    }

    private static ByteArrayOutputStream concatPdf(List<byte[]> listPdfsBytes) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        PDFMergerUtility merger = new PDFMergerUtility();
        for (byte[] pdf : listPdfsBytes) {
            if (pdf!=null&&pdf.length > 0) {
                merger.addSource(new ByteArrayInputStream(pdf));
            }
        }
        merger.setDestinationStream(byteArrayOutputStream);
        merger.mergeDocuments(MemoryUsageSetting.setupTempFileOnly());
        return byteArrayOutputStream;
    }

}
