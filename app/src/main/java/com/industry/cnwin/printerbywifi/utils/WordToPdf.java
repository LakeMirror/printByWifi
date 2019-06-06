package com.industry.cnwin.printerbywifi.utils;

import android.content.Context;


import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Stack;


public class WordToPdf{

    public static void docx2Html(Context mContext, XWPFDocument document, String outPutFile) throws IOException {
        String fileOutName = outPutFile;
        long startTime = System.currentTimeMillis();
        XHTMLOptions options = XHTMLOptions.create().indent(4);
        // 导出图片
        File imageFolder = new File(mContext.getFilesDir() + File.separator + "images");
        options.setExtractor(new FileImageExtractor(imageFolder));
        // URI resolver  word的html中图片的目录路径
        options.URIResolver(new BasicURIResolver("images"));
        File outFile = new File(fileOutName);
        outFile.getParentFile().mkdirs();
        OutputStream out = new FileOutputStream(outFile);
        XHTMLConverter.getInstance().convert(document, out, options);
        System.out.println("Generate " + fileOutName + " with " + (System.currentTimeMillis() - startTime) + " ms.");
    }
}
