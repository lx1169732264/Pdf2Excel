package org.example;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;
import org.apache.pdfbox.util.Matrix;

import java.io.File;

public class Pdf2Excel3 {

    public static void main(String[] args) {
        try {
            PDDocument document =  PDDocument.load(new File("C:\\Users\\lx\\Desktop\\source.pdf"));
            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setStartPage(1); // 设置要处理的页面范围
            stripper.setEndPage(document.getNumberOfPages());
            stripper.setSortByPosition(true);
            stripper.setWordSeparator("|||");
            String text = stripper.getText(document);
            System.out.println(text);
            // 将提取的文本写入Excel文件
            // 可以使用Apache POI等库来创建Excel文件并将文本写入其中
            document.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
