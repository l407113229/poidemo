package poi;

import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.itext.extension.font.ITextFontRegistry;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.awt.*;
import java.io.InputStream;
import java.io.OutputStream;


public class DocxToPDFConverter extends Converter {


    public DocxToPDFConverter(InputStream inStream,
                              OutputStream outStream,
                              boolean showMessages,
                              boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }


    @Override
    public void convert() throws Exception {
        //loading();

        PdfOptions options = PdfOptions.create();
        XWPFDocument document = new XWPFDocument(inStream);


        //支持中文字体
        options.fontProvider(new ITextFontRegistry() {
            @Override
            public Font getFont(String familyName, String encoding, float size, int style, Color color) {
                try {
                    Resource fileSource = new ClassPathResource("SimSun.ttf");
                    String path = fileSource.getFile().getAbsolutePath();

                    BaseFont bfChinese = BaseFont.createFont(path, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font fontChinese = new Font(bfChinese, size, style, color);
                    if (familyName != null) {
                        fontChinese.setFamily(familyName);
                    }
                    return fontChinese;
                } catch (Throwable e) {
                    e.printStackTrace();
                    return ITextFontRegistry.getRegistry().getFont(familyName, encoding, size, style, color);
                }
            }


        });


        //processing();
        PdfConverter.getInstance().convert(document, outStream, options);


        finished();
    }


}