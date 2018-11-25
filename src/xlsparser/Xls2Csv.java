package xlsparser;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import org.w3c.dom.Document;
import org.xml.sax.SAXException;

/**
 * @author rspotti
 * Static Class Xls2Csv, convert Excel xls file (xml formatted) into csv file
 *
 */
class Xls2Csv {
	private static char fieldSep='|';
	
	/**
	 * private method to generate XSLT file used to convert xls file to csv
	 * generate a temporary file deleted at the end of execution
	 * 
	 * DEFAULT XSL:
	 * <xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	 * xmlns:qq="urn:schemas-microsoft-com:office:spreadsheet" version="1.0">
	 * <xsl:output method="xml" indent="yes" omit-xml-declaration="yes" />
	 * <xsl:variable name="separator" select="'|'"/>
	 * <xsl:variable name="rowseparator" >
	 *    <xsl:text>&#x0A;</xsl:text> 
	 * </xsl:variable> 
	 * <xsl:template match="/" name="riga" >
	 *   <xsl:for-each select="Cell">
	 *     <xsl:value-of select="."/>
	 *     <xsl:value-of select="$separator" />
     *   </xsl:for-each>	
     * </xsl:template>
     *
     * <xsl:template match="/">
	 *   <xsl:for-each select="Workbook/Worksheet/Table/Row">
	 *     <xsl:call-template name="riga" />
	 *     <xsl:value-of select="$rowseparator" />
     *   </xsl:for-each>	
     * </xsl:template>
     * </xsl:stylesheet>

	 *  
	 * @return a File object containing the temporary xslt model
	 */
	private static File writeXSL()
    {	
    	try{
    		File temp = File.createTempFile("tempXSL", ".xsl");
    		temp.deleteOnExit();
	        BufferedWriter bw = new BufferedWriter(new FileWriter(temp));   
    	    bw.write("<xsl:stylesheet\n xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\"\n");
    	    bw.write("xmlns:qq=\"urn:schemas-microsoft-com:office:spreadsheet\" version=\"1.0\">\n");
    	    bw.write("<xsl:output method=\"xml\" indent=\"yes\" omit-xml-declaration=\"yes\" />\n");
    	    bw.write("<xsl:variable name=\"separator\" select=\"'"+fieldSep+"'\"/>\n");
    	    bw.write("<xsl:variable name=\"rowseparator\" >\n");
    	    bw.write("<xsl:text>&#x0A;</xsl:text>\n");  
    	    bw.write("</xsl:variable>\n"); 
    	    bw.write("<xsl:template match=\"/\" name=\"riga\" >\n");
    	    bw.write("	<xsl:for-each select=\"Cell\">\n");
    	    bw.write("		<xsl:value-of select=\".\"/>\n");
    	    bw.write("		<xsl:value-of select=\"$separator\" />\n");
    	    bw.write("	</xsl:for-each>\n</xsl:template>\n");
    	    bw.write("<xsl:template match=\"/\">\n");
    	    bw.write("	<xsl:for-each select=\"Workbook/Worksheet/Table/Row\">\n");
    	    bw.write("		<xsl:call-template name=\"riga\" />\n");
    	    bw.write("		<xsl:value-of select=\"$rowseparator\" />\n");
    	    bw.write("	</xsl:for-each>\n</xsl:template>\n</xsl:stylesheet>\n");
    	    bw.close();
    	    return temp;
    	}catch(IOException e){
    		e.printStackTrace();
    		return null;
    	}
    }
	
	/**
	 * Set the default field separator as the supplied parameter field 
	 *  if not specified it uses separator is '|'.
	 * @param field
	 */
	public static void setFieldSeparator(char field) {
		fieldSep=field;
	}
	
	/**
	 * 
	 * @param xlsFile
	 * @param csvFile
	 * @param field
	 * 
	 * @return true if succeed, false otherwise.
	 */
	public static boolean transform(String xlsFile, String csvFile,char field) {
		setFieldSeparator(field);
		return transform(xlsFile, csvFile);
	}
	
	/**
	 * Run conversion of the file xlsFile to csvFile setting and using default field separator 
	 * @param xlsFile
	 * @param csvFile
	 * 
	 * @return true if succeed, false otherwise.
	 */
	public static boolean transform(String xlsFile, String csvFile) {
		File stylesheet=writeXSL();
		boolean result=false;
		if (null != stylesheet) {
			result=transform(xlsFile, csvFile, stylesheet);
		}
		return result;
	}
	
	/**
	 * Run conversion of the file xlsFile to csvFile, allow to specify a different xslt file.
	 * 
	 * @param xlsFile
	 * @param csvFile
	 * @param xslFile
	 * 
	 * @return true if succeed, false otherwise.
	 */
	public static boolean transform(String xlsFile, String csvFile, String xslFile) {	
		File stylesheet = new File(xslFile);
		return transform(xlsFile, csvFile, stylesheet);
	}
			
	/**
	 * Convert xlsFile to csvFile using stylesheet as xslt file
	 * this is internally used by all the pubblic method.
	 * 
	 * @param xlsFile
	 * @param csvFile
	 * @param stylesheet
	 * 
	 * @return true if succeed, false otherwise.
	 */
	private static boolean transform(String xlsFile, String csvFile, File stylesheet) {
  
    	File xmlSource = new File(xlsFile);
    	
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        try {
        	DocumentBuilder builder = factory.newDocumentBuilder();
        	Document document = builder.parse(xmlSource);

        	StreamSource stylesource = new StreamSource(stylesheet);
        	Transformer transformer = TransformerFactory.newInstance()
                .newTransformer(stylesource);
        	Source source = new DOMSource(document);
        	Result outputTarget = new StreamResult(new File(csvFile));
        	//
        	// Conversion
			transformer.transform(source, outputTarget);
		} catch (TransformerException | ParserConfigurationException | SAXException | IOException e) {
			e.printStackTrace();
			return false;
		}
        return true;
    }
}
// Xls2Csv