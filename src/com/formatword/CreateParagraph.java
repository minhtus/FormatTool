/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.formatword;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.impl.schema.SchemaTypeImpl;
import org.apache.xmlbeans.impl.schema.SchemaTypeSystemImpl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTRowImpl;

/**
 *
 * @author Phi
 */
public class CreateParagraph {

    public static void main(String[] args) throws IOException {
//        XWPFDocument document = new XWPFDocument();
//        FileOutputStream fos = new FileOutputStream(new File("noinaycoanh.docx"));
//        XWPFParagraph paragraph = document.createParagraph();
//        XWPFRun run = paragraph.createRun();
//        run.setText("Nguoi theo huong  hoa may mu giang loi lang suong khoi phoi pha dua buoc ai xa roi");
//        document.write(fos);
//        fos.close();
//        System.out.println("Write document successfully");
        FileInputStream fis = new FileInputStream(new File("twocolumn.docx"));
        XWPFDocument document = new XWPFDocument(fis);
        XWPFParagraph tmpParagraph = document.getParagraphs().get(0);

        for (int i = 0; i < 100; i++) {
            XWPFRun tmpRun = tmpParagraph.createRun();
            tmpRun.setText("LALALALAALALAAAA");
            tmpRun.setFontSize(18);
        }
        XWPFTable table = document.createTable();
        table.setInsideHBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "");
        table.setInsideVBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "");
        XWPFTableRow row = table.createRow();
        row.createCell().setText("Tac gia");
        row.createCell().setText("Son Tung MTP");
        row.createCell().setText("Tac gia");
        row.createCell().setText("Den");
        row = table.createRow();
        row.createCell().setText("Tac pham");
        row.createCell().setText("Lac troi");
        row.createCell().setText("Tac pham");
        row.createCell().setText("Abc");
        row = table.createRow();
        row.createCell().setText("Danh gia");
        row.createCell().setText("Khong gi sanh bang");
        row.createCell().setText("Danh gia");
        row.createCell().setText("Khong gi sanh bang 123");
        document.write(new FileOutputStream(new File("poi.docx")));
    }
}
