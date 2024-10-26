package com.future.rocket.ice;

import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.SaveFormat;
import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;
import com.lowagie.text.pdf.BaseFont;
import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;
import org.junit.Test;
import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import java.awt.*;
import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

public class TestIceBlueMain {

    @Test
    public void test0() {
        System.out.println("hi");
    }

    @Test
    public void test1() {

        //许可证加载
        //loadElicFile();

        //创建一个Document实例
        Document document = new Document();

        ParagraphStyle style = new ParagraphStyle(document);
        style.setName("Heading1");
        style.getCharacterFormat().setBold(true);
        style.getCharacterFormat().setFontSize(16);
        style.getCharacterFormat().setFontName("Arial");
        style.getCharacterFormat().setTextColor(Color.BLUE);
        document.getStyles().add(style);

        ParagraphStyle style2 = new ParagraphStyle(document);
        style2.setName("Heading2");
        style2.getCharacterFormat().setFontSize(10);
        style2.getCharacterFormat().setFontName("Arial");
        style2.getCharacterFormat().setTextColor(Color.RED);
        document.getStyles().add(style2);

        //添加一个节
        Section section = document.addSection();

        //添加合同标题
        Paragraph titleParagraph = section.addParagraph();
        titleParagraph.appendText("租房合同\n");
        titleParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        titleParagraph.applyStyle("Heading1");

        //添加合同编号
        Paragraph contractNumberParagraph = section.addParagraph();
        contractNumberParagraph.appendText("合同编号: HT-2024-001\n");
        contractNumberParagraph.applyStyle("Normal");

        //添加合同双方信息
        Paragraph partiesParagraph = section.addParagraph();
        partiesParagraph.appendText("(甲方)恒源建材有限公司\n");
        partiesParagraph.appendText("(乙方)东兴水泥厂\n");
        partiesParagraph.appendText("签署日期: 2024年10月23日\n\n");
        partiesParagraph.applyStyle("Normal");

        //添加合同正文
        Paragraph contentParagraph = section.addParagraph();
        contentParagraph.appendText("合同正文: \n");
        contentParagraph.appendText("1. 租房租金为每月20000元整 \n");
        contentParagraph.appendText("2. 有效期从2024-10-23日起,2025-10-22日止 \n");
        contentParagraph.appendText("3. 房屋用途为住宿,不能进行二次转租 \n");
        contentParagraph.applyStyle("Normal");

        //添加合同签署
        Paragraph signatureParagraph = section.addParagraph();
        signatureParagraph.appendText("\n甲方代表签字: ________________ \n");
        signatureParagraph.appendText("乙方代表签字: ________________ \n");
        signatureParagraph.applyStyle("Heading2");

        String wordFilePath = "src/main/resources/contract.doc";
        document.saveToFile(wordFilePath, FileFormat.Docx_2013);
        System.out.println("合同 Word 文档生成 ==> " + wordFilePath);

        String pdfFilePath = "src/main/resources/contract.pdf";
        document.saveToFile(pdfFilePath, FileFormat.PDF);
        System.out.println("合同 Pdf 文档生成 ==> " + pdfFilePath);
    }

    @Test
    public void test2() throws Exception {
        com.aspose.words.Document document = new com.aspose.words.Document();
        DocumentBuilder builder = new DocumentBuilder(document);

        //设置标题样式
        builder.getFont().setBold(true);
        builder.getFont().setSize(18);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.writeln("租赁合同\n");

        //恢复正常字体大小和样式
        builder.getFont().clearFormatting();
        builder.getFont().setSize(12);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);

        //添加合同编号
        builder.writeln("合同编号: HT-202410-001\n");

        //添加甲乙双方
        builder.writeln("(甲方)吴小峰");
        builder.writeln("(乙方)王小二");
        builder.writeln("签署日期: 2024-10-23日\n");

        //插入合同条款
        builder.writeln("合同正文:");
        builder.writeln("1. 租金每月 800 元整");
        builder.writeln("2. 每月1-5号缴纳");
        builder.writeln("3. 不可二次转租\n");

        //合同金额
        builder.writeln("合同金额: ¥800/月\n");

        builder.writeln("\n甲方代表签字:___吴小峰___");
        builder.writeln("乙方代表签字:___王小二___");

        //保存word 文档
        String wordFile = "src/main/resources/contractX.doc";
        document.save(wordFile);
        System.out.println("合同 Word 文档生成 ==> " + wordFile);

        //保存pdf 文档
        String pdfFile = "src/main/resources/contractX.pdf";
        document.save(pdfFile, SaveFormat.PDF);
        System.out.println("合同 PDF 文档生成 ==> " + pdfFile);
    }

    @Test
    public void test3() {
        String htmlFilePath = "src/main/resources/contract.html";
        String pdfPath = "src/main/resources/contractH.pdf";

        try {
            String htmlContent = loadHtmlFromFile(htmlFilePath);
            generatePdfFromHtml(htmlContent, pdfPath);
            System.out.println("Pdf 文件生成 ==>" + pdfPath);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    @Test
    public void test4() {
        // 创建合同信息
        Map<String, String> contractDetails = new TreeMap<>();
        contractDetails.put("Date", "2024-10-23");
        contractDetails.put("Solar Term", "Frost's Descent");
        contractDetails.put("Blessing Recipient", "XueLin TongXue");
        contractDetails.put("Blessing Message", "Happy Frost's Descent to XueLin (2024-10-23)! From Yao ^_^.");
        contractDetails.put("From", "Yao");

        // 生成合同信息字符串，使用统一的换行符处理
        StringBuilder contractInfo = new StringBuilder();
        contractDetails.forEach((key, value) -> contractInfo.append(key).append(": ").append(value).append(" | "));

        // 确保去除最后一个多余的换行符
        //String contractInfoString = contractInfo.toString().trim();

        String contractInfoString = "Happy Frost's Descent to XueLin (2024-10-23)! From Yao ^_^.";

        String filePath = "src/main/resources/contractQRCode.png";
        generateQRCode(contractInfoString, filePath);
        System.out.println("二维码生成成功 ==> " + filePath);
    }

    private void generateQRCode(String data, String filePath) {
        try {
            int width = 400;  // 增加二维码的尺寸以适应更多内容
            int height = 400;

            // 配置二维码参数，设置字符集为UTF-8
            Map<EncodeHintType, Object> hints = new HashMap<>();
            hints.put(EncodeHintType.CHARACTER_SET, "UTF-8");  // 设置字符编码
            hints.put(EncodeHintType.MARGIN, 1);  // 边距设置较小，避免空间浪费
            hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);  // 设置高容错级别

            // 生成二维码
            BitMatrix bitMatrix = new MultiFormatWriter().encode(data, BarcodeFormat.QR_CODE, width, height, hints);
            Path path = Paths.get(filePath);

            // 输出为图片
            MatrixToImageWriter.writeToPath(bitMatrix, "PNG", path);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    //html文件生成pdf
    private void generatePdfFromHtml(String htmlContent, String pdfPath) throws Exception {
        ITextRenderer renderer = new ITextRenderer();

        // 加载 SimSun 中文字体
        ITextFontResolver fontResolver = renderer.getFontResolver();
        fontResolver.addFont("src/main/resources/SimSun.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        // 设置文档内容
        renderer.setDocumentFromString(htmlContent);
        renderer.layout();

        // 输出 PDF 文件
        try (FileOutputStream fos = new FileOutputStream(new File(pdfPath))) {
            renderer.createPDF(fos);
        }
    }

    //获取html文件内容
    private String loadHtmlFromFile(String htmlFilePath) throws IOException {
        StringBuilder contentBuilder = new StringBuilder();
        try(BufferedReader br = new BufferedReader(new FileReader(htmlFilePath))) {
            String currentLine;
            while ((currentLine = br.readLine()) != null) {
                contentBuilder.append(currentLine).append("\n");
            }
        }
        return contentBuilder.toString();
    }

    //加载冰蓝License文件
    private void loadElicFile() {
        try (InputStream inputStream = TestIceBlueMain.class.getClassLoader().getResourceAsStream("license.elic.xml")) {
            com.spire.license.LicenseProvider.setLicenseFile(inputStream);
        } catch (Exception e) {
            System.out.println("冰蓝软件初始化失败 ==> " + e);
        }
    }
}