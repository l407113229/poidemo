package poi;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import lombok.Data;

import java.io.File;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * @author liuyasong
 * @date 2022/5/25 12:14
 */
public class Test {

    // PDF 格式
    private static final int wdFormatPDF = 17;

    public static void main(String[] args) throws Exception {
        URL filePath = Test.class.getClassLoader().getResource("Test.docx");
        assert filePath != null;
        String targetPath = "D:/target.docx";
        String desPath = "D:/jacob.pdf";

        List<KeyInventoryDetail> details = new ArrayList<>();
        KeyInventoryDetail detail = new KeyInventoryDetail();

        details.add(detail);
        details.add(detail);
        details.add(detail);
        details.add(detail);

        LoopRowTableRenderPolicy hackLoopTableRenderPolicy = new LoopRowTableRenderPolicy();
        Configure config = Configure.builder().bind("keyInventoryDetail", hackLoopTableRenderPolicy)
                .build();

        XWPFTemplate template = XWPFTemplate.compile(filePath.getPath(), config).render(new HashMap<String, Object>() {
            {
                put("jobNo", "001");
                put("keyInventoryDetail", details);

            }
        });
        template.writeToFile(targetPath);
        wordToPDF(targetPath, desPath);

    }


    static void wordToPDF(String docxPath, String pdfPath) {

        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", new Variant(false));
            Dispatch docs = app.getProperty("Documents").toDispatch();


            doc = Dispatch.call(docs, "Open", docxPath).toDispatch();
            File tofile = new File(pdfPath);
            if (tofile.exists()) {
                tofile.delete();
            }
            Dispatch.call(doc, "SaveAs", pdfPath, wdFormatPDF);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        } finally {
            Dispatch.call(doc, "Close", false);
            if (app != null) {
                app.invoke("Quit", new Variant[]{});
            }
        }
        //结束后关闭进程
        ComThread.Release();
    }


    @Data
    static class KeyInventoryDetail {
        private String no;
        private String warehouseNo;
        private String cargoSpace;
        private String materialCode;
        private String materialDes;
        private String batchNo;
        private String cargoSpaceCnt;
        private String actualCnt;
        private String unit;
        private String storeLv;
        private String storeMethod;
        private String packageMethod;
        private String longevity;
        private String expiredDate;

        public KeyInventoryDetail() {
            no = "1";
            warehouseNo = "SH001";
            cargoSpace = "N1";
            materialCode = "materialCode";
            materialDes = "materialDes";
            batchNo = "batchNo";
            cargoSpaceCnt = "cargoSpaceCnt";
            actualCnt = "actualCnt";
            unit = "unit";
            storeLv = "storeLv";
            storeMethod = "storeMethod";
            packageMethod = "packageMethod";
            longevity = "longevity";
            expiredDate = "expiredDate";
        }

    }
}
