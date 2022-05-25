package poi;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import lombok.Data;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * @author liuyasong
 * @date 2022/5/25 12:14
 */
public class Test {

    public static void main(String[] args) throws IOException {
        URL filePath = Test.class.getClassLoader().getResource("Test.docx");
        assert filePath != null;
        String targetPath = "D:/target.docx";
        String desPath = "D:/openOffice.pdf";

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

        Word2Pdf(targetPath, desPath);
    }


    /**
     * 将word格式的文件转换为pdf格式
     *
     * @param srcPath 原地址
     * @param desPath 目标地址
     * @throws IOException 异常
     */
    public static void Word2Pdf(String srcPath, String desPath) throws IOException {
        // 源文件目录
        File inputFile = new File(srcPath);
        if (!inputFile.exists()) {
            System.out.println("源文件不存在！");
            return;
        }
        // 输出文件目录
        File outputFile = new File(desPath);
        if (!outputFile.getParentFile().exists()) {
            outputFile.getParentFile().exists();
        }
        // 调用openoffice服务线程
        String command = "C:/Program Files (x86)/OpenOffice 4/program/soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";
        Process p = Runtime.getRuntime().exec(command);

        // 连接openoffice服务
        OpenOfficeConnection connection = new SocketOpenOfficeConnection(
                "127.0.0.1", 8100);
        connection.connect();

        // 转换word到pdf
        DocumentConverter converter = new OpenOfficeDocumentConverter(
                connection);
        converter.convert(inputFile, outputFile);

        // 关闭连接
        connection.disconnect();

        // 关闭进程
        p.destroy();
        System.out.println("转换完成！");
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
