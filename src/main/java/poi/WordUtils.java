package poi;

import lombok.extern.slf4j.Slf4j;
import org.docx4j.Docx4J;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.org.apache.poi.util.IOUtils;

import java.io.File;
import java.io.FileOutputStream;

@Slf4j
public class WordUtils {

    ///**
    // * docx文档转换为PDF
    // * @param body 文档
    // * @param response 响应给前端
    // * @return pdf 输出流
    // * @throws Exception 可能为Docx4JException, FileNotFoundException, IOException等
    // */
    //public static void convertDocxToPdf(byte[] body , HttpServletResponse response) throws Exception {
    //    response.setContentType("application/pdf");
    //    File docxFile = FileUtil.byteToFile(body, UUID.randomUUID().toString() + ".docx");
    //    try {
    //        WordprocessingMLPackage mlPackage = WordprocessingMLPackage.load(docxFile);
    //        setFontMapper(mlPackage);
    //        Docx4J.toPDF(mlPackage, response.getOutputStream());
    //    }catch (Exception e){
    //        e.printStackTrace();
    //        log.error("docx文档转换为PDF失败");
    //    }
    //   FileUtil.deleteTempFile(docxFile);
    //}


    /**
     * docx文档转换为PDF
     *
     * @param pdfPath PDF文档存储路径
     * @throws Exception 可能为Docx4JException, FileNotFoundException, IOException等
     */
    public static void convertDocxToPdf(String docxPath, String pdfPath) throws Exception {

        FileOutputStream fileOutputStream = null;
        try {
            File file = new File(docxPath);
            fileOutputStream = new FileOutputStream(pdfPath);
            WordprocessingMLPackage mlPackage = WordprocessingMLPackage.load(file);
            setFontMapper(mlPackage);
            Docx4J.toPDF(mlPackage, new FileOutputStream(pdfPath));
        } catch (Exception e) {
            e.printStackTrace();
            log.error("docx文档转换为PDF失败");
        } finally {
            IOUtils.closeQuietly(fileOutputStream);
        }
    }

    private static void setFontMapper(WordprocessingMLPackage mlPackage) throws Exception {
        Mapper fontMapper = new IdentityPlusMapper();
        fontMapper.put("隶书", PhysicalFonts.get("LiSu"));
        fontMapper.put("宋体", PhysicalFonts.get("SimSun"));
        fontMapper.put("微软雅黑", PhysicalFonts.get("Microsoft Yahei"));
        fontMapper.put("黑体", PhysicalFonts.get("SimHei"));
        fontMapper.put("楷体", PhysicalFonts.get("KaiTi"));
        fontMapper.put("新宋体", PhysicalFonts.get("NSimSun"));
        fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));
        fontMapper.put("华文仿宋", PhysicalFonts.get("STFangsong"));
        fontMapper.put("宋体扩展", PhysicalFonts.get("simsun-extB"));
        fontMapper.put("仿宋", PhysicalFonts.get("FangSong"));
        fontMapper.put("仿宋_GB2312", PhysicalFonts.get("FangSong_GB2312"));
        fontMapper.put("幼圆", PhysicalFonts.get("YouYuan"));
        fontMapper.put("华文宋体", PhysicalFonts.get("STSong"));
        fontMapper.put("华文中宋", PhysicalFonts.get("STZhongsong"));

        mlPackage.setFontMapper(fontMapper);
    }
}