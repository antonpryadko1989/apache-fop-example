package com.visoft;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.common.PDStream;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

import javax.imageio.ImageIO;
import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

import static org.apache.commons.io.FileUtils.writeByteArrayToFile;


public class PDFMerge {

    private static List<byte[]> bytes = new ArrayList<>();
    private static String pdfPath = "C:\\Users\\user\\Downloads\\old.pdf";
    private static String bigPdfPath = "C:\\Users\\user\\Downloads\\1.pdf";
    private static String imagePdfPath1 = "C:\\Users\\user\\Downloads\\1.jpg";
    private static String imagePdfPath2 = "C:\\Users\\user\\Downloads\\1.png";
    private static String imagePdfPath3 = "C:\\Users\\user\\Downloads\\1.jpeg";
    private static String fakeImagePdfPath = "C:\\Users\\user\\Downloads\\fake.jpg";
    private static String newPath = "C:\\Users\\user\\Downloads\\new.pdf";

    private static String cv = "C:\\Users\\user\\Downloads\\CV.pdf";
    private static String im1 = "C:\\Users\\user\\Downloads\\im1.png";
    private static String im2 = "C:\\Users\\user\\Downloads\\im2.png";
    private static String im3 = "C:\\Users\\user\\Downloads\\im3.png";


    private static byte[] pdfContent;
    private static byte[] bigPdfContent;
    private static byte[] imagePdfContent1;
    private static byte[] imagePdfContent2;
    private static byte[] imagePdfContent3;

    private static byte[] cvContent;
    private static byte[] im1Content;
    private static byte[] im2Content;
    private static byte[] im3Content;

    static {try {pdfContent = Files.readAllBytes(new File(pdfPath).toPath());} catch (IOException e) {e.printStackTrace();}}
    static {try {bigPdfContent = Files.readAllBytes(new File(bigPdfPath).toPath());} catch (IOException e) {e.printStackTrace();}}
    static {try {imagePdfContent1 = Files.readAllBytes(new File(imagePdfPath1).toPath());} catch (IOException e) {e.printStackTrace();}}
    static {try {imagePdfContent2 = Files.readAllBytes(new File(imagePdfPath2).toPath());} catch (IOException e) {e.printStackTrace();}}
    static {try {imagePdfContent3 = Files.readAllBytes(new File(imagePdfPath3).toPath());} catch (IOException e) {e.printStackTrace();}}

    static {try {cvContent = Files.readAllBytes(new File(cv).toPath());} catch (IOException e) {e.printStackTrace();}}
    static {try {im1Content = Files.readAllBytes(new File(im1).toPath());} catch (IOException e) {e.printStackTrace();}}
    static {try {im2Content = Files.readAllBytes(new File(im2).toPath());} catch (IOException e) {e.printStackTrace();}}
    static {try {im3Content = Files.readAllBytes(new File(im3).toPath());} catch (IOException e) {e.printStackTrace();}}

    public static void main(String[] args) throws IOException {
        byte[] pdfBytes = null;

        bytes.add(cvContent);
        pdfBytes = convertImageToPdf(im1Content);
        if(pdfBytes!=null ){bytes.add(pdfBytes);}

        pdfBytes = convertImageToPdf(im2Content);
        if(pdfBytes!=null ){bytes.add(pdfBytes);}

        pdfBytes = convertImageToPdf(im3Content);
        if(pdfBytes!=null ){bytes.add(pdfBytes);}
        bytes.add(pdfContent);
        pdfBytes = convertImageToPdf(imagePdfContent2);
        if(pdfBytes!=null ){
            bytes.add(pdfBytes);
        }
        bytes.add(bigPdfContent);
        pdfBytes = convertImageToPdf(imagePdfContent3);
        if(pdfBytes!=null ){
            bytes.add(pdfBytes);
        }
        pdfBytes = convertImageToPdf(imagePdfContent1);
        if(pdfBytes!=null ){
            bytes.add(pdfBytes);
        }
        bytes.add(pdfContent);
        bytes.add(bigPdfContent);
        bytes.add(bigPdfContent);
        writeData(newPath, concatPdf(bytes));
    }

    private static byte[] concatPdf(List<byte[]> listPdfsBytes) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        PDFMergerUtility merger = new PDFMergerUtility();
        for (byte[] pdf : listPdfsBytes) {
            if (pdf!=null&&pdf.length > 0) {
                merger.addSource(new ByteArrayInputStream(pdf));
            }
        }
        merger.setDestinationStream(byteArrayOutputStream);
        merger.mergeDocuments(MemoryUsageSetting.setupTempFileOnly());
        return byteArrayOutputStream.toByteArray();
    }

//    private static byte[] imagesToPdf(byte[] image) {
//        try {
//            PDDocument doc = new PDDocument();
//            InputStream in = null;
//            PDPageContentStream contentStream = null;
//            final ByteArrayOutputStream baos = new ByteArrayOutputStream();
//            in = new ByteArrayInputStream(image);
//            BufferedImage bi = ImageIO.read(in);
//            float width = bi.getWidth();
//            float height = bi.getHeight();
//            PDPage page = new PDPage(new PDRectangle(width, height));
//            doc.addPage(page);
//            PDImageXObject pdImage = PDImageXObject.createFromByteArray(doc, image, null);
//            contentStream = new PDPageContentStream(doc, page, PDPageContentStream.AppendMode.APPEND, true, true);
//            float scale = 1f;
//            contentStream.drawImage(pdImage, 0, 0, pdImage.getWidth() * scale, pdImage.getHeight() * scale);
//            System.out.println();
//            return baos.toByteArray();
//        }catch (Exception e){
//
//        }
//        return result;
//    }


    private static byte[] m1(byte[] imageBytes) throws IOException {
        byte[] result = null;
        PDPage page = new PDPage(
                new PDRectangle(
                        ImageIO.read(
                                new ByteArrayInputStream(imageBytes)
                        ).getWidth(),
                        ImageIO.read(
                                new ByteArrayInputStream(imageBytes)
                        ).getHeight()
                )
        );
        PDDocument doc = new PDDocument();
        PDPageContentStream contentStream = new PDPageContentStream(doc, page, PDPageContentStream.AppendMode.APPEND, true, true);
        try{
            doc.addPage(page);
            PDImageXObject pdImage = PDImageXObject.createFromByteArray(doc, imageBytes, null);
            float scale = 1f;
            contentStream.drawImage(pdImage, 0, 0, pdImage.getWidth() * scale, pdImage.getHeight() * scale);
            result = new PDStream(doc).createInputStream().toString().getBytes();

        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }

    private static byte[] convertImageToPdf(byte[] imageBytes){
        byte[] result = null;
        try{
            Document document;
            ByteArrayOutputStream stream = new ByteArrayOutputStream();

            Image img = Image.getInstance(imageBytes);
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
            e.printStackTrace();
        }
        return result;

    }

    private static void writeData(String path, byte[] data){
        try {
            writeByteArrayToFile(new File(path), data);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
