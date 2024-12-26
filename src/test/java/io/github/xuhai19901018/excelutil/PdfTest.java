package io.github.xuhai19901018.excelutil;


import com.aspose.cells.License;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;


public class PdfTest {
    @Test
    public void t1() {

        try {

            InputStream is = this.getClass().getClassLoader().getResourceAsStream("license.xml");//这个文件应该是类似于密码验证(证书？)，用于获得去除水印的权限

            License aposeLic = new License();

            aposeLic.setLicense(is);

            Workbook wbk = new Workbook("D:\\home\\aaa.xlsx");

            wbk.save("D:\\home\\aaa.pdf", SaveFormat.PDF);

        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }


    }
    @Test
    public void t2() {

        try {



        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }
//    @Test
//    public void t4() {
//
//        try {
//
//
////            ExcelUtils.addValue("logo", "D:\\home\\日油.png");
//
//            // 输出Excel
//            ExcelUtils.exportPdf("D:\\home\\s5w.xlsx", new FileOutputStream(new File("D:\\home\\t4.pdf")), PaperSizeType.PAPER_A_4, PageOrientationType.LANDSCAPE,"C:\\Windows\\Fonts");
//        } catch (Exception e) {
//            // TODO Auto-generated catch block
//            e.printStackTrace();
//        }
//
//    }

}
