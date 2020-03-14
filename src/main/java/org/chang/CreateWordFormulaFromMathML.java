package org.chang;

import java.io.*;

import org.apache.poi.xwpf.usermodel.*;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTR;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.xmlbeans.XmlCursor;

/*
needs the full ooxml-schemas-1.3.jar as mentioned in https://poi.apache.org/faq.html#faq-N10025
*/

public class CreateWordFormulaFromMathML {

    static File stylesheet = new File("MML2OMML.XSL");
    static TransformerFactory tFactory = TransformerFactory.newInstance();
    static StreamSource stylesource = new StreamSource(stylesheet);

    static CTOMath getOMML(String mathML) throws Exception {
        Transformer transformer = tFactory.newTransformer(stylesource);

        StringReader stringreader = new StringReader(mathML);
        StreamSource source = new StreamSource(stringreader);

        StringWriter stringwriter = new StringWriter();
        StreamResult result = new StreamResult(stringwriter);
        transformer.transform(source, result);

        String ooML = stringwriter.toString();
        stringwriter.close();

        CTOMathPara ctOMathPara = CTOMathPara.Factory.parse(ooML);
        CTOMath ctOMath = ctOMathPara.getOMathArray(0);

        //for making this to work with Office 2007 Word also, special font settings are necessary
        XmlCursor xmlcursor = ctOMath.newCursor();
        while (xmlcursor.hasNextToken()) {
            XmlCursor.TokenType tokentype = xmlcursor.toNextToken();
            if (tokentype.isStart()) {
                if (xmlcursor.getObject() instanceof CTR) {
                    CTR cTR = (CTR) xmlcursor.getObject();
                    cTR.addNewRPr2().addNewRFonts().setAscii("Cambria Math");
                    cTR.getRPr2().getRFonts().setHAnsi("Cambria Math");
                }
            }
        }

        return ctOMath;
    }

    public static void main(String[] args) throws Exception {

        XWPFDocument document = new XWPFDocument();

        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("The Pythagorean theorem: ");

        String mathML =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                        + "<mrow>"
                        + "<msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup>"
                        + "</mrow>"
                        + "</math>";

        CTOMath ctOMath = getOMML(mathML);
        System.out.println(ctOMath);

        CTP ctp = paragraph.getCTP();
        ctp.setOMathArray(new CTOMath[]{ctOMath});

        paragraph = document.createParagraph();
        run = paragraph.createRun();
        run.setText("The Quadratic Formula: ");

        mathML =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                        + "<mrow>"
                        + "<mi>x</mi><mo>=</mo><mfrac><mrow><mrow><mo>-</mo><mi>b</mi></mrow><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>-</mo><mrow><mn>4</mn><mo>⁢</mo><mi>a</mi><mo>⁢</mo><mi>c</mi></mrow></mrow></msqrt></mrow><mrow><mn>2</mn><mo>⁢</mo><mi>a</mi></mrow></mfrac>"
                        + "</mrow>"
                        + "</math>";

        ctOMath = getOMML(mathML);
        System.out.println(ctOMath);

        ctp = paragraph.getCTP();
        ctp.setOMathArray(new CTOMath[]{ctOMath});

        document.write(new FileOutputStream("CreateWordFormulaFromMathML.docx"));
        document.close();

    }
}