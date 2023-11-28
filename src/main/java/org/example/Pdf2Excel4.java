package org.example;

import jxl.CellView;
import jxl.Workbook;
import jxl.format.VerticalAlignment;
import jxl.write.*;
import net.sf.json.JSONArray;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.junit.Test;
import technology.tabula.CommandLineApp;

import java.io.FileOutputStream;
import java.io.OutputStream;

import static jxl.format.Alignment.CENTRE;

public class Pdf2Excel4 {

    @Test
    public void test() {
        String s = extractFromPdf("C:\\Users\\lx\\Desktop\\source.pdf");
        //解析tabula读取pdf表格,将返回的数据转成jsonArray
        JSONArray jsonArray = new JSONArray();
        jsonArray.add(s);

        try (OutputStream o = new FileOutputStream("C:\\Users\\lx\\Desktop\\target.xlsx")) {
            WritableWorkbook workbook = Workbook.createWorkbook(o);
            WritableSheet sheet1 = workbook.createSheet("sheet1", 0);

            JSONArray jsonPage = jsonArray.getJSONArray(0), dataArr = new JSONArray();
            for (int i = 0; i < jsonPage.size(); i++) {
                dataArr.addAll(jsonPage.getJSONObject(i).getJSONArray("data"));
            }

            //获取每页中的data
            for (int k = 0; k < dataArr.size(); k++) {
                //遍历data中的每一条,也就是单元格中的每一行
                JSONArray dataD = dataArr.getJSONArray(k);
                //通过下标获取每个单元格的数据
                int idx = 0;
                WritableCellFormat format = new WritableCellFormat();
                format.setAlignment(CENTRE);
                format.setVerticalAlignment(VerticalAlignment.CENTRE);
                while (idx < dataD.size()) {
                    if (0 == Double.parseDouble(dataD.getJSONObject(idx).get("top").toString().replaceAll("\r", ""))) {
                        Blank blank = new jxl.write.Blank(sheet1.getCell(idx, k));
                        blank.setCellFormat(format);
                        sheet1.addCell(blank);
                        sheet1.mergeCells(idx, k - 1, idx, k);
                    } else {
                        String text = dataD.getJSONObject(idx).get("text").toString().replaceAll("\r", "");
                        Label label = new Label(idx, k, text);

                        label.setCellFormat(format);
                        sheet1.addCell(label);
                    }


                    idx++;
                }
            }

            //宽度自适应
            for (int i = 0; i < sheet1.getColumns(); i++) {
                CellView cellView = new CellView();
                cellView.setAutosize(true);
                sheet1.setColumnView(i,cellView);
            }


            workbook.write();
            workbook.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private String extractFromPdf(String pdfPath) {
        try {
            //-f导出格式,默认CSV  (一定要大写)  -p 指导出哪页,all是所有
            String[] argsa = new String[]{"-f=JSON", "-p=all", pdfPath, "-l"};
            CommandLineParser parser = new DefaultParser();
            CommandLine cmd = parser.parse(CommandLineApp.buildOptions(), argsa);
            StringBuilder stringBuilder = new StringBuilder();
            new CommandLineApp(stringBuilder, cmd).extractTables(cmd);
            return stringBuilder.toString();
        } catch (Exception e) {
            //todo
            throw new RuntimeException();
        }
    }
}
