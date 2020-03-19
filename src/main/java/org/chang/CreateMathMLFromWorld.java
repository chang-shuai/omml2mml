package org.chang;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.w3c.dom.Node;

import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import java.io.*;
import java.net.URL;
import java.util.*;

public class CreateMathMLFromWorld {
    private static File stylesheet = new File("OMML2MML.XSL");
    private static TransformerFactory tFactory = TransformerFactory.newInstance();
    private static StreamSource stylesource = new StreamSource(stylesheet);


    /**
     * 获取MathML
     * @param ctomath
     * @return
     * @throws Exception
     */
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

        mathML = mathML.replaceAll("xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"", "");
        mathML = mathML.replaceAll("xmlns:mml", "xmlns");
        mathML = mathML.replaceAll("mml:", "");

        return mathML;
    }

    /**
     * 返回公式的集合
     * @param document
     * @return
     */
    public static Map<Integer,String> mml2Html(XWPFDocument document){
        Map<Integer,String> result = new HashMap<>();
        try{

            //storing the found MathML in a AllayList of strings
            List<String> mathMLList = new ArrayList<String>(16);

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

            for (int i = 0; i < mathMLList.size(); i++) {
                // 替换特殊符号(由于页面上无法直接展示特殊符号,所以需要进行替换,将特殊符号替换为html可以认识的标签(https://www.cnblogs.com/xinlvtian/p/8646683.html))
                String s = mathMLList.get(i)
                        .replaceAll("±", "&#x00B1;")
                        .replaceAll("∑","&sum;");
                s = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" + s + "</math>";
                result.put(i,s);
            }
            return result;
        }catch (Exception e){
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 获取所有的公式
     * @param xwpfDocument
     * @return
     */
    public static Map<Integer,String> getFormulaMap(XWPFDocument xwpfDocument) {
        Map<Integer, String> result = new HashMap<>();
        // 获取到公式的Map集合
        Map<Integer, String> mml2Html = mml2Html(xwpfDocument);
        Set<Map.Entry<Integer, String>> entries = mml2Html.entrySet();
        // 遍历所有段落,获取所有包含公式的段落
        List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
        int j = 0;
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph xwpfParagraph = paragraphs.get(i);
            CTP ctp = xwpfParagraph.getCTP();
            String xmlText = ctp.xmlText();
            if(xmlText.contains("<m:oMath>")){
                StringBuilder sb = new StringBuilder();
                sb.append(xwpfParagraph.getParagraphText());
                sb.append(mml2Html.get(j++));
                result.put(i,sb.toString());
            }

        }
        return result;
    }

    public static String convertMML2Latex(String mml) throws FileNotFoundException {
        mml = mml.substring(mml.indexOf("?>")+2, mml.length()); //去掉xml的头节点
        URIResolver r = new URIResolver(){  //设置xls依赖文件的路径
            @Override
            public Source resolve(String href, String base) {
                InputStream inputStream = null;

                try {
                    inputStream = new FileInputStream("mmltex.xsl");
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
                return new StreamSource(inputStream);

            }
        };
        String latex = xslConvert(mml, "mmltex.xsl", null);
        if(latex != null && latex.length() > 1){
            latex = latex.substring(1, latex.length() - 1);
        }
        return latex;
    }
    public static String xslConvert(String s, String xslpath, URIResolver uriResolver) throws FileNotFoundException {
        TransformerFactory tFac = TransformerFactory.newInstance();
        if(uriResolver != null)  tFac.setURIResolver(uriResolver);
        StreamSource xslSource = new StreamSource(new FileInputStream(xslpath));
        StringWriter writer = new StringWriter();
        try {
            Transformer t = tFac.newTransformer(xslSource);
            Source source = new StreamSource(new StringReader(s));
            Result result = new StreamResult(writer);
            t.transform(source, result);
        } catch (TransformerException e) {
        }
        return writer.getBuffer().toString();
    }

    public static String convertOMML2MML(String xml) throws FileNotFoundException {
        String result = xslConvert(xml, "OMML2MML.XSL", null);
        return result;
    }
    public static void main(String[] args) {
        String path = "docx/content.docx";
        try(InputStream inputStream = CreateMathMLFromWorld.class.getClassLoader().getResourceAsStream(path)) {
            XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
            Map<Integer, String> formulaMap = getFormulaMap(xwpfDocument);

        } catch (IOException e) {
            e.printStackTrace();
        }

        // 这个就能获取到所有公式了,Integer表示的是第几个公式,String表示公式转化后的mathML,借助mathJax可以在页面上进行展示
//        String mml = convertOMML2MML(omml);
    }
}
