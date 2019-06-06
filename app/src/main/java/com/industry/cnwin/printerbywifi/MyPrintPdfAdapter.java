package com.industry.cnwin.printerbywifi;

import android.os.Build;
import android.os.Bundle;
import android.os.CancellationSignal;
import android.os.ParcelFileDescriptor;
import android.print.PageRange;
import android.print.PrintAttributes;
import android.print.PrintDocumentAdapter;
import android.print.PrintDocumentInfo;
import android.support.annotation.RequiresApi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

@RequiresApi(api = Build.VERSION_CODES.KITKAT)
public class MyPrintPdfAdapter extends PrintDocumentAdapter {
    private String mFilePath;
 
    public MyPrintPdfAdapter(String file) {
        this.mFilePath = file;
    }
 
    @RequiresApi(api = Build.VERSION_CODES.KITKAT)
    @Override
    public void onLayout(PrintAttributes oldAttributes, PrintAttributes newAttributes,
                         CancellationSignal cancellationSignal,
                         LayoutResultCallback callback, Bundle extras) {
        if (cancellationSignal.isCanceled()) {
            callback.onLayoutCancelled();
            return;
        }
        PrintDocumentInfo info = new PrintDocumentInfo.Builder("name")
                .setContentType(PrintDocumentInfo.CONTENT_TYPE_DOCUMENT)
                .build();
        callback.onLayoutFinished(info, true);
    }
 
    @RequiresApi(api = Build.VERSION_CODES.KITKAT)
    @Override
    public void onWrite(PageRange[] pages, ParcelFileDescriptor destination,
                        CancellationSignal cancellationSignal,
                        WriteResultCallback callback) {
        InputStream input = null;
        OutputStream output = null;
 
        try {
 
            input = new FileInputStream(mFilePath);
            output = new FileOutputStream(destination.getFileDescriptor());
 
            byte[] buf = new byte[1024];
            int bytesRead;
 
            while ((bytesRead = input.read(buf)) > 0) {
                output.write(buf, 0, bytesRead);
            }
 
            callback.onWriteFinished(new PageRange[]{PageRange.ALL_PAGES});
 
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                output.close();           
            } catch (IOException e) {
                e.printStackTrace();
            }
            
            try {
                input.close();                
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}