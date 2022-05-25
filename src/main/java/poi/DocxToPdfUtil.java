package poi;

import com.itextpdf.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Objects;

/**
 * @author liuyasong
 * @date 2022/5/25 22:53
 */
public class DocxToPdfUtil {

    /**
     * docx格式word转换为html
     *
     * @param fileNamePath   docx文件路径
     * @param outPutFilePath html输出文件路径
     * @throws TransformerException
     * @throws IOException
     * @throws ParserConfigurationException
     */
    public static void docx2Html(String fileNamePath, String outPutFilePath) throws TransformerException, IOException, ParserConfigurationException {
        long startTime = System.currentTimeMillis();
        XWPFDocument document = new XWPFDocument(new FileInputStream(fileNamePath));
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        //1、强转中文格式类型，解决中文消失问题
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                if (run.getFontFamily() == "Calibri") {
                    run.setFontFamily("SimHei");
                }
                run.setFontFamily("SimSun", XWPFRun.FontCharRange.ascii);
            }
        }
        XHTMLOptions options = XHTMLOptions.create().indent(4);
        File outFile = new File(outPutFilePath);
        outFile.getParentFile().mkdirs();
        OutputStreamWriter outputStreamWriter = new OutputStreamWriter(new FileOutputStream(outFile), StandardCharsets.UTF_8);
        XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
        xhtmlConverter.convert(document, outputStreamWriter, options);

        //String html = FileUtil.readFileToString(outPutFilePath, "html");

        File file = new File(outPutFilePath);
        FileInputStream fileInputStream = new FileInputStream(file);
        byte[] data = new byte[fileInputStream.available()];
        fileInputStream.read(data);
        fileInputStream.close();

        String html = new String(data, StandardCharsets.UTF_8);

        OutputStreamWriter writer = new OutputStreamWriter(new FileOutputStream(outPutFilePath), StandardCharsets.UTF_8);
        //2、添加标准html头部，解决中文乱码问题
        html = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">" + html;
        int i = html.indexOf("<head>");
        StringBuilder buffer = new StringBuilder(html);
        html = buffer.insert(i + 6, "<style type=\"text/css\">\n" +
                "    *\n" +
                "    {\n" +
                "        padding-left: 20pt;\n" +
                "        padding-right: -20pt;\n" +
                "    }\n" +
                "</style>").toString();
        writer.write(html);
        writer.flush();
        writer.close();
        String rootPath = System.getProperty("user.dir") + "/src/main/resources";
        html2pdf(html, "D:/html.pdf", rootPath);
        System.out.println("Generate " + outPutFilePath + " with " + (System.currentTimeMillis() - startTime) + " ms.");
    }

    /**
     * docx格式word转换为html
     *
     * @param html    html文件
     * @param pdfName pdf文件名
     * @param fontDir 指定字体文件夹路径
     */

    public static void html2pdf(String html, String pdfName, String fontDir) {
        try {
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            ITextRenderer renderer = new ITextRenderer();
            ITextFontResolver fontResolver = (ITextFontResolver) renderer.getSharedContext().getFontResolver();
            //遍历添加中文字体库
            File f = new File(fontDir);
            if (f.isDirectory()) {
                File[] files = f.listFiles((dir, name) -> {
                    String lower = name.toLowerCase();
                    return lower.endsWith(".otf") || lower.endsWith(".ttf") || lower.endsWith(".ttc");
                });
                for (int i = 0; i < Objects.requireNonNull(files).length; i++) {
                    fontResolver.addFont(files[i].getAbsolutePath(), BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                }
            }
            //添加字体库结束
            renderer.setDocumentFromString(html);
            renderer.layout();
            renderer.createPDF(os);
            renderer.finishPDF();
            byte[] buff = os.toByteArray();
            //保存到磁盘上
            File file = new File(pdfName);
            //创建文件字节输出流对象
            FileOutputStream outputStream = new FileOutputStream(file);
            BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(outputStream);
            bufferedOutputStream.write(buff);
            bufferedOutputStream.close();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
