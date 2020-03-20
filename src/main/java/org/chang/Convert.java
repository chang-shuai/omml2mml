package org.chang;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.List;

public class Convert {

    public static void docxToHtml(String path, String file) throws Throwable {
        InputStream input = new FileInputStream(path + file);
        XWPFDocument document = new XWPFDocument(input);
        XHTMLOptions options = XHTMLOptions.create();
        File imageFolder = new File(path);
        // 设置图片生成路径
        options.setExtractor(new FileImageExtractor(imageFolder));
        // 设置图片链接，去掉后使用相对路径
        // options.URIResolver(new FileURIResolver(imageFolder));
        OutputStream out = new FileOutputStream(new File(path + "index.html"));
        XHTMLConverter.getInstance().convert(document, out, options);
    }

    /**
     * Convert doc to HTML
     * @param path
     * @param file
     * @throws Throwable
     */
    public static void docToHtml(String path, String file) throws Throwable {
        InputStream input = new FileInputStream(path + file);
        HWPFDocument wordDocument = new HWPFDocument(input);
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            @Override
            public String savePicture(byte[] bytes, PictureType pictureType, String s, float v, float v1) {
                return s;
            }

        });
        wordToHtmlConverter.processDocument(wordDocument);
        List<?> pics = wordDocument.getPicturesTable().getAllPictures();
        if (pics != null) {
            for (int i = 0; i < pics.size(); i++) {
                Picture pic = (Picture) pics.get(i);
                try {
                    pic.writeImageContent(new FileOutputStream(path + pic.suggestFullFileName()));
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
            }
        }
        Document htmlDocument = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);

        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        outStream.close();
        String content = new String(outStream.toByteArray());
        writeFile(content, path + "index.html", "UTF-8");
    }

    /**
     * 生成 Html 文件
     * @param content
     * @param path
     * @param encode
     */
    public static void writeFile(String content, String path, String encode) {
        FileOutputStream fos = null;
        BufferedWriter bw = null;
        try {
            File file = new File(path);
            fos = new FileOutputStream(file);
            bw = new BufferedWriter(new OutputStreamWriter(fos, encode));
            bw.write(content);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (bw != null) {
                    bw.close();
                }
                if (fos != null) {
                    fos.close();
                }
            } catch (IOException ie) {
                ie.printStackTrace();
            }
        }
    }

    public static void main(String[] args) throws Throwable {
        final String path = "D:/test/";
        String path1 = Convert.class.getClassLoader().getResource("docx/yuwen.docx").getPath();
        System.out.println(path1);
        final String file = "504b5d8f-5105-43ee-9229-938c8426acfd_word_content.doc";
        // DocToHtml.convert(path, file);
        docToHtml(path, file);
    }
}
