package com.visoft.services;

import java.io.*;
import java.nio.file.Paths;
import java.util.*;

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
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;
import org.xml.sax.SAXException;

@Service
public class FOPPdf {

	@Value("${templates.repository}")
	private String templatesRepository;

	@Value("${xml.start}")
	private String xmlStart;

	@Value("${xml.end}")
	private String xmlEnd;

	@Autowired
	private Input2OutputService resultService;

	StreamingResponseBody getPDFFile(TemplateDTO template) {
		File f =  Paths.get(templatesRepository,  template.getTemplateName()).toFile();
		if(f.isDirectory()){
			throw new PathValidationException("error", "It's directory");
		}
		if(!f.exists()){
			throw new PathValidationException("error", "File not exist");
		}
		try {
			return convertPDF(template);
		} catch (IOException | TransformerException | SAXException e) {
			e.printStackTrace();
		}
		throw new PathValidationException("error", "Something get wrong");
	}

	private StreamingResponseBody convertPDF(TemplateDTO template)
			throws IOException, TransformerException, SAXException{
		File xsltFile = Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()).toFile();
		JSONObject jsonObject = new JSONObject(template);
		String str = xmlStart  + XML.toString(jsonObject) + xmlEnd;
		StreamSource xmlSource = new StreamSource(new ByteArrayInputStream(str.getBytes()));
		FopFactory fopFactory = FopFactory.newInstance(Paths.get(templatesRepository, "/userconfig.xml").toFile());
		FOUserAgent foUserAgent = fopFactory.newFOUserAgent();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		try {
			Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, foUserAgent, out);
			TransformerFactory factory = TransformerFactory.newInstance();
			Transformer transformer = factory.newTransformer(new StreamSource(xsltFile));
			Result res = new SAXResult(fop.getDefaultHandler());
			transformer.transform(xmlSource, res);
		} finally {
			out.close();
		}
		InputStream inputStream = new ByteArrayInputStream(out.toByteArray());
		return resultService.getOutput(inputStream);
	}
}
