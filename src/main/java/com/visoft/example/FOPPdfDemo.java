package com.visoft.example;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import javax.xml.transform.Result;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import org.apache.fop.apps.FOPException;
import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.Fop;
import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.MimeConstants;
import org.xml.sax.SAXException;

/**
 * @author vlad
 *
 */
public class FOPPdfDemo {


	///fgljkgfd

	public static final String RESOURCES_DIR;
	public static final String OUTPUT_DIR;

	static {
		RESOURCES_DIR = "src//main//resources//";
		OUTPUT_DIR = "src//main//resources//output//";
	}

	public static void main(String[] args) {

		FOPPdfDemo fOPPdfDemo = new FOPPdfDemo();

		try {
			fOPPdfDemo.convertToPDF("//template-HE.xsl", "//Employees-HE.xml", "//employee-HE.pdf");
			fOPPdfDemo.convertToPDF("//template-RU.xsl", "//Employees-RU.xml", "//employee-RU.pdf");
		} catch (IOException | TransformerException | SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * Method that will convert the given XML to PDF
	 * 
	 * @throws IOException
	 * @throws TransformerException
	 * @throws SAXException
	 */
	public void convertToPDF(final String template, final String dataXML, final String outputPDF)
			throws IOException, TransformerException, SAXException {
		// the XSL FO file
		File xsltFile = new File(RESOURCES_DIR + template);
		// the XML file which provides the input
		StreamSource xmlSource = new StreamSource(
				new File(RESOURCES_DIR + dataXML));
		// create an instance of fop factory
		FopFactory fopFactory = FopFactory
				.newInstance(new File(RESOURCES_DIR + "//userconfig.xml"));
		// a user agent is needed for transformation
		FOUserAgent foUserAgent = fopFactory.newFOUserAgent();
		// Setup output
		OutputStream out;
		out = new java.io.FileOutputStream(OUTPUT_DIR + outputPDF);

		try {
			// Construct fop with desired output format
			Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, foUserAgent,
					out);

			// Setup XSLT
			TransformerFactory factory = TransformerFactory.newInstance();
			Transformer transformer = factory
					.newTransformer(new StreamSource(xsltFile));

			// Resulting SAX events (the generated FO) must be piped through to
			// FOP
			Result res = new SAXResult(fop.getDefaultHandler());

			// Start XSLT transformation and FOP processing
			// That's where the XML is first transformed to XSL-FO and then
			// PDF is created
			transformer.transform(xmlSource, res);
		} finally {
			out.close();
		}
	}

	/**
	 * This method will convert the given XML to XSL-FO
	 * 
	 * @throws IOException
	 * @throws FOPException
	 * @throws TransformerException
	 */
	public void convertToFO()
			throws IOException, FOPException, TransformerException {
		// the XSL FO file
		File xsltFile = new File(RESOURCES_DIR + "//template-HE.xsl");

		// the XML file which provides the input
		StreamSource xmlSource = new StreamSource(
				new File(RESOURCES_DIR + "//Employees.xml"));

		OutputStream out;

		out = new java.io.FileOutputStream(OUTPUT_DIR + "temp.fo");

		try {
			// Setup XSLT
			TransformerFactory factory = TransformerFactory.newInstance();
			Transformer transformer = factory
					.newTransformer(new StreamSource(xsltFile));

			Result res = new StreamResult(out);

			// Start XSLT transformation and FOP processing
			transformer.transform(xmlSource, res);

			// Start XSLT transformation and FOP processing
			// That's where the XML is first transformed to XSL-FO and then
			// PDF is created
			transformer.transform(xmlSource, res);
		} finally {
			out.close();
		}
	}

}
