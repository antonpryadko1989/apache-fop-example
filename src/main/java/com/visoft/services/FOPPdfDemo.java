package com.visoft.services;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Date;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamSource;

import com.visoft.exceptions.PathValidationException;
import com.visoft.templates.entity.TemplateDTO;
import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.Fop;
import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.MimeConstants;
import org.json.JSONObject;
import org.json.XML;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.xml.sax.SAXException;

@Service
public class FOPPdfDemo {

	@Value("${templates.repository}")
	private String templatesRepository;

//	private static final String RESOURCES = "E:/WORK/TEMPLATES/";

	private static final String RESOURCES_DIR = "src/main/resources/";

	String getPDFUrl(TemplateDTO template) {

		File f =  Paths.get(templatesRepository, template.getProjectId(),  template.getTemplateName()).toFile();
		if(f.isDirectory()){
			throw new PathValidationException("error", "It's directory");
		}
		if(!f.exists()){
			throw new PathValidationException("error", "File not exist");
		}

		try {
			return convertPDF(template);
		} catch (IOException | TransformerException | SAXException | ParserConfigurationException e) {
			e.printStackTrace();
		}
		throw new PathValidationException("error", "Something get wrong");
	}

	private String convertPDF(TemplateDTO template)
			throws IOException, TransformerException, SAXException, ParserConfigurationException {
		File xsltFile = Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()).toFile();
		JSONObject jsonObject = new JSONObject(template);
		String str = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><root>"  + XML.toString(jsonObject) + "</root>";
		Path path = Files.createTempFile(template.getProjectId(), ".xml");
		File file = path.toFile();
		Files.write(path, str.getBytes(StandardCharsets.UTF_8));
		file.deleteOnExit();
		String tempPath = file.getAbsolutePath();
		StreamSource xmlSource = new StreamSource(new File(tempPath));

		FopFactory fopFactory = FopFactory.newInstance(Paths.get(RESOURCES_DIR, "/userconfig.xml").toFile());
		FOUserAgent foUserAgent = fopFactory.newFOUserAgent();
		String outFile = Paths.get(templatesRepository, template.getProjectId(), new Date().getTime() + template.getOutPutName()).toString();
		OutputStream out;
		out = new java.io.FileOutputStream(outFile);
		try {
			Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, foUserAgent, out);
			TransformerFactory factory = TransformerFactory.newInstance();
			Transformer transformer = factory.newTransformer(new StreamSource(xsltFile));
			Result res = new SAXResult(fop.getDefaultHandler());
			transformer.transform(xmlSource, res);
		} finally {
			out.close();
			Files.delete(Paths.get(tempPath));
		}
		return outFile;
	}
}
