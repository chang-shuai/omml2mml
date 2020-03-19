package org.chang;

import java.io.*;
import org.apache.poi.xwpf.usermodel.*;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;

import org.w3c.dom.Node;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamSource;
import javax.xml.transform.stream.StreamResult;

import java.awt.Desktop;

import java.util.List;
import java.util.ArrayList;

/*
needs the full ooxml-schemas-1.3.jar as mentioned in https://poi.apache.org/faq.html#faq-N10025
*/

public class WordReadFormulas {

    static File stylesheet = new File("OMML2MML.XSL");
    static TransformerFactory tFactory = TransformerFactory.newInstance();
    static StreamSource stylesource = new StreamSource(stylesheet);

    static String getMathML(CTOMath ctomath) throws Exception {
        Transformer transformer = tFactory.newTransformer(stylesource);

        Node node = ctomath.getDomNode();

        DOMSource source = new DOMSource(node);
        StringWriter stringwriter = new StringWriter();
        StreamResult result = new StreamResult(stringwriter);
        transformer.setOutputProperty("omit-xml-declaration", "yes");
        transformer.transform(source, result);

        String mathML = stringwriter.toString();
        stringwriter.close();

        //The native OMML2MML.XSL transforms OMML into MathML as XML having special name spaces.
        //We don't need this since we want using the MathML in HTML, not in XML.
        //So ideally we should changing the OMML2MML.XSL to not do so.
        //But to take this example as simple as possible, we are using replace to get rid of the XML specialities.
        mathML = mathML.replaceAll("xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"", "");
        mathML = mathML.replaceAll("xmlns:mml", "xmlns");
        mathML = mathML.replaceAll("mml:", "");

        return mathML;
    }

    public static void main(String[] args) throws Exception {
        InputStream inputStream = WordReadFormulas.class.getClassLoader().getResourceAsStream("docx/数学试题.docx");
        XWPFDocument document = new XWPFDocument(inputStream);

        //storing the found MathML in a AllayList of strings
        List<String> mathMLList = new ArrayList<String>();

        //getting the formulas out of all body elements
        for (IBodyElement ibodyelement : document.getBodyElements()) {
            if (ibodyelement.getElementType().equals(BodyElementType.PARAGRAPH)) {
                XWPFParagraph paragraph = (XWPFParagraph)ibodyelement;
                for (CTOMath ctomath : paragraph.getCTP().getOMathList()) {
                    mathMLList.add(getMathML(ctomath));
                }
                for (CTOMathPara ctomathpara : paragraph.getCTP().getOMathParaList()) {
                    for (CTOMath ctomath : ctomathpara.getOMathList()) {
                        mathMLList.add(getMathML(ctomath));
                    }
                }
            } else if (ibodyelement.getElementType().equals(BodyElementType.TABLE)) {
                XWPFTable table = (XWPFTable)ibodyelement;
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            for (CTOMath ctomath : paragraph.getCTP().getOMathList()) {
                                mathMLList.add(getMathML(ctomath));
                            }
                            for (CTOMathPara ctomathpara : paragraph.getCTP().getOMathParaList()) {
                                for (CTOMath ctomath : ctomathpara.getOMathList()) {
                                    mathMLList.add(getMathML(ctomath));
                                }
                            }
                        }
                    }
                }
            }
        }

        document.close();

        //creating a sample HTML file
        String encoding = "UTF-8";
        FileOutputStream fos = new FileOutputStream("result.html");
        OutputStreamWriter writer = new OutputStreamWriter(fos, encoding);
        writer.write("<!DOCTYPE html>\n");
        writer.write("<html lang=\"en\">");
        writer.write("<head>");
        writer.write("<meta charset=\"utf-8\"/>");

        //using MathJax for helping all browsers to interpret MathML
        writer.write("<script type=\"text/javascript\"");
//        writer.write(" async src=\"https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.1/MathJax.js?config=MML_CHTML\"");
        writer.write(" async src=\"http://dev.bcbook.cn/MathJax-master/MathJax.js?config=MML_CHTML\"");
        writer.write(">");
        writer.write("</script>");

        writer.write("</head>");
        writer.write("<body>");
        writer.write("<p>Following formulas was found in Word document: </p>");

        int i = 1;
        for (String mathML : mathMLList) {
            writer.write("<p>Formula" + i++ + ":</p>");
            writer.write(mathML);
            writer.write("<p/>");
        }

        writer.write("</body>");
        writer.write("</html>");
        writer.close();

        Desktop.getDesktop().browse(new File("result.html").toURI());

    }
}