package com.industry.cnwin.printerbywifi;

import android.content.Context;
import android.net.Uri;
import android.util.Log;

import com.industry.cnwin.printerbywifi.custom.CustomXWPFDocument;
import com.industry.cnwin.printerbywifi.utils.WordToPdf;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.ref.WeakReference;
import java.math.BigInteger;
import java.util.Objects;

public class WordGenerate {

    private WeakReference<Context> mContex;
    private static volatile WordGenerate instance;
    private CustomXWPFDocument customXWPFDocument;
    private OutputStream os;


    private WordGenerate(Context context) {
        mContex = new WeakReference<>(context);
    }

    public static WordGenerate newInstance(Context context) {
        if (instance == null) {
            synchronized (WordGenerate.class) {
                if (instance == null) {
                    instance = new WordGenerate(context);
                }

            }
        }
        return instance;
    }

    private File word;
    private CustomXWPFDocument xwpfDocument;

    public WordGenerate newWordFile(String name) {
        try {
            String path = mContex.get().getFilesDir() + File.separator + name + ".docx";
            word = new File(path);
            os = new FileOutputStream(word);
            xwpfDocument = new CustomXWPFDocument();
            PackagePart pp = xwpfDocument.createStyles().getPackagePart();
            xwpfDocument.getPackagePart().addRelationship(pp.getPartName(), TargetMode.INTERNAL, pp.getContentType());
            xwpfDocument.getDocument().getBody().addNewSectPr();
        } catch (Exception e) {

        }
        return this;
    }

    public WordGenerate insertTitle(String title) {
        if (xwpfDocument == null) {
            throw new IllegalStateException("Invole Method newWordFile() First");
        }

        createWord(ParagraphAlignment.CENTER, 20, true, title);
        return this;

    }

    public WordGenerate insertWord(ParagraphAlignment align, int fontSize, boolean bold, String word) {
        if (xwpfDocument == null) {
            throw new IllegalStateException("Invole Method newWordFile() First");
        }
        createWord(align, fontSize, bold, word);
        return this;
    }

    public WordGenerate insertPicture(FileInputStream inputStream) {
        try {
            if (xwpfDocument == null) {
                throw new IllegalStateException("Invole Method newWordFile() First");
            }
            byte[] ba = new byte[inputStream.available()];
            inputStream.read(ba);
            ByteArrayInputStream byteInputStream = new ByteArrayInputStream(ba);
            XWPFParagraph picture = xwpfDocument.createParagraph();
//            添加图片
            xwpfDocument.addPictureData(byteInputStream, CustomXWPFDocument.PICTURE_TYPE_JPEG);
            xwpfDocument.createPicture(xwpfDocument.getAllPictures().size() - 1, 100, 100, picture);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return this;

    }

    public WordGenerate insertExcel(String[][] excelData) {
        if (xwpfDocument == null) {
            throw new IllegalStateException("Invole Method newWordFile() First");
        }
        if (excelData.length == 0) {
            Log.e("PrintByWifi", "excelData's size is 0");
            return this;
        }
        XWPFTable table = xwpfDocument.createTable(excelData.length, excelData[0].length);
        CTTbl ttbl = table.getCTTbl();
        CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
        CTTblBorders borders = tblPr.isSetTblBorders() ? tblPr.getTblBorders(): tblPr.addNewTblBorders();

        ttbl.addNewTblGrid();
        for (int i = 0; i < table.getRows().size(); i++) {
            XWPFTableRow row = table.getRow(i);
            for (int j = 0; j < row.getTableCells().size(); j++) {
                XWPFTableCell cell = row.getCell(j);
                CTTc cttc = cell.getCTTc();
                CTTcPr cellPr = cttc.isSetTcPr() ? cttc.getTcPr() : cttc.addNewTcPr();
                cellPr.addNewTcW();
                cellPr.getTcW();
                cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                if (i == 0) {
                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
                } else {
                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.RIGHT);
                }
                CTTblWidth tblWidth1 = cellPr.isSetTcW() ? cellPr.getTcW() : cellPr.addNewTcW();
                tblWidth1.setW(new BigInteger("800"));
                tblWidth1.setType(STTblWidth.DXA);
                cell.setText(excelData[i][j]);
                XWPFParagraph paragraph = cell.addParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText(excelData[i][j]);

                ((XWPFParagraph)cell.getBodyElements().get(0)).addRun(run);
            }
        }
        return this;
    }

    public String generate() {
        String dst = "";
        try {
            if (xwpfDocument == null) {
                throw new IllegalStateException("Invole Method newWordFile() First");
            }
            xwpfDocument.write(os);
            os.flush();
            os.close();
            dst =  mContex.get().getFilesDir() + File.separator + "test.html";
//            WordToPdf.docxToPdf(xwpfDocument, dst);
            WordToPdf.docx2Html(mContex.get(),  xwpfDocument, dst);

            return dst;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return dst;
    }

    /**
     * 生成文字
     *
     * @param align    文字的位置 CENTER LEFT RIGHT
     * @param fontSize 文字的大小
     * @param bold     文字是否加粗
     * @param message  文字的内容
     */
    private void createWord(ParagraphAlignment align, int fontSize, boolean bold, String message) {
        XWPFParagraph titleParagraph = xwpfDocument.createParagraph();
        titleParagraph.setAlignment(align);
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText(message);
        titleRun.setFontSize(fontSize);
        titleRun.setFontFamily("宋体");
        titleRun.setBold(bold);
    }

}
