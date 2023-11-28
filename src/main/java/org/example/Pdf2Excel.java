//package org.example;
//
//import java.io.FileOutputStream;
//import java.io.InputStream;
//
//import com.aspose.pdf.*;
//
//public class Pdf2Excel {
//
//    public static void main(String[] args) {
//        String sourceFile = "C:\\Users\\lx\\Desktop\\source.pdf",
//                targetFile = "C:\\Users\\lx\\Desktop\\target.xlsx";
//
//        try (InputStream is = Pdf2Excel.class.getResourceAsStream("/license.xml");//授权文件
//             Document doc = new Document(sourceFile);//加载源文件数据
//             FileOutputStream os = new FileOutputStream(targetFile)) {
//
//            //调用授权方法
//            License license = new License();
//            license.setLicense(is);
//            doc.save(os, SaveFormat.Excel);//设置转换文件类型并转换
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//}