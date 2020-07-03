package com.example.asposedemo;

import com.aspose.cells.License;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

@SpringBootTest
class AsposeDemoApplicationTests {

    /**
     * 获取license
     *
     * @return
     */
    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = Test.class.getClassLoader().getResourceAsStream("\\license.xml");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 支持DOC, DOCX, OOXML, RTF, HTML, OpenDocument, PDF, EPUB, XPS, SWF等相互转换<br>
     *
     * @param args
     */
    public static void main(String[] args) {
        // 验证License
        if (!getLicense()) {
            return;
        }
        try {
            long old = System.currentTimeMillis();
            // 转换html
            ExcelConvertToHtml("D:\\\\BaiduNetdiskDownload\\\\aspose-demo\\\\src\\\\main\\\\resources\\\\static\\\\aaa.xlsx",
                    "D:\\\\BaiduNetdiskDownload\\\\aspose-demo\\\\src\\\\main\\\\resources\\\\static\\\\aaa.html");

            // 转换csv
            ExcelConvertToCSV("D:\\\\BaiduNetdiskDownload\\\\aspose-demo\\\\src\\\\main\\\\resources\\\\static\\\\aaa.xlsx",
                    "D:\\\\BaiduNetdiskDownload\\\\aspose-demo\\\\src\\\\main\\\\resources\\\\static\\\\aaa.csv");

            // 转换pdf
            Workbook wb = new Workbook("D:\\BaiduNetdiskDownload\\aspose-demo\\src\\main\\resources\\static\\aaa.xlsx");// 原始excel路径
            File pdfFile = new File("D:\\BaiduNetdiskDownload\\aspose-demo\\src\\main\\resources\\static\\aaa.pdf");// 输出路径
            FileOutputStream fileOS = new FileOutputStream(pdfFile);
            wb.save(fileOS, SaveFormat.PDF);

            long now = System.currentTimeMillis();
            System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Excel 2 csv转换类
     * @param sourceFilePath 传入路径
     * @param csvFilePath   输出路径
     * @throws Exception 异常
     */
    public static void ExcelConvertToCSV(String sourceFilePath, String csvFilePath)
            throws Exception {
        com.aspose.cells.Workbook excel = null;
        excel = new com.aspose.cells.Workbook(sourceFilePath);
        excel.save(csvFilePath, com.aspose.cells.SaveFormat.CSV);
    }

    /**
     * Excel 2 HTML转换类
     * @param sourceFilePath 传入路径
     * @param htmlFilePath   输出路径
     * @throws Exception 异常
     */
    public static void ExcelConvertToHtml(String sourceFilePath, String htmlFilePath)
            throws Exception {
        com.aspose.cells.LoadOptions loadOption = null;
        com.aspose.cells.Workbook excel = null;
        if (sourceFilePath != null
                && !sourceFilePath.isEmpty()
                && sourceFilePath
                .substring(sourceFilePath.lastIndexOf("."))
                .toLowerCase() == ".csv") {
            loadOption = new com.aspose.cells.TxtLoadOptions(
                    com.aspose.cells.LoadFormat.AUTO);
        }
        if (loadOption != null) {
            excel = new com.aspose.cells.Workbook(sourceFilePath, loadOption);
        } else {
            excel = new com.aspose.cells.Workbook(sourceFilePath);
        }
        excel.save(htmlFilePath, com.aspose.cells.SaveFormat.HTML);
    }

}
