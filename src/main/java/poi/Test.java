package poi;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import lombok.Data;

import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * @author liuyasong
 * @date 2022/5/25 12:14
 */
public class Test {

    public static void main(String[] args) throws Exception {
        URL filePath = Test.class.getClassLoader().getResource("Test.docx");
        assert filePath != null;
        String targetPath = "D:/target.docx";
        String tempPath = "D:/temp.html";
        String desPath = "D:/html.pdf";

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

        DocxToPdfUtil.docx2Html(targetPath,tempPath);
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
