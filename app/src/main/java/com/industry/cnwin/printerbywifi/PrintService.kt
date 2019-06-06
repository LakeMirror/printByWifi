package com.industry.cnwin.printerbywifi

import android.annotation.TargetApi
import android.content.Context
import android.graphics.Bitmap
import android.print.PrintAttributes
import android.print.PrintManager
import android.support.annotation.RequiresApi
import android.support.v4.print.PrintHelper
import android.util.Log
import android.webkit.WebView
import android.webkit.WebViewClient
import com.google.gson.Gson
import android.print.PrintDocumentAdapter
import android.annotation.SuppressLint
import android.os.*
import android.widget.Toast
import android.print.PageRange
import com.android.dx.stock.ProxyBuilder
import java.io.File
import java.io.IOException
import java.lang.reflect.InvocationHandler


class PrintService() {

    companion object {
        @Volatile
        private var INSTANCE: PrintService? = null

        /*获取单例*/
        val instance: PrintService
            get() {
                if (INSTANCE == null) {
                    synchronized(PrintService::class.java) {
                        if (INSTANCE == null) {
                            INSTANCE = PrintService()
                        }
                    }
                }
                return INSTANCE!!
            }
    }

    fun printBitmap(context: Context, bitmap: Bitmap) {
        Log.i("PrintManager", "START PRINT")
        val printHelper = PrintHelper(context)
        printHelper.scaleMode = PrintHelper.SCALE_MODE_FIT
        printHelper.printBitmap("test-print", bitmap)
        Log.i("PrintManager", "END PRINT")
    }

    fun printExcelByWebView(context: Context, url: String, data: Array<Array<String>>) {
        printHTML(context, url, data, false)
    }

    fun printPDFByWebView(context: Context, url: String, data: Array<Array<String>>) {
       printHTML(context, url, data, true)
    }

    fun printWebView(context: Context, webView: WebView) {
        if (Build.VERSION.SDK_INT > 19) {
            createWebPrintJob(context, webView)
        } else {
            Toast.makeText(context, "当前 Android 版本低于 4.4", Toast.LENGTH_SHORT).show()
        }
    }

    @RequiresApi(Build.VERSION_CODES.KITKAT)
    fun printWord(context: Context, path: String) {
//        val dst = "${context.filesDir}${File.separator}printPDF.pdf"
//        WordToPdf.docxToPdf(XWPFDocument( FileInputStream(path)), dst)
//        printPDFFile(context, dst, WebView(context))

        printHTML(context, path, null, false)
    }

    private fun printHTML(context: Context, url: String, data: Array<Array<String>>?, isPdf: Boolean) {
        var webView: WebView? = WebView(context)
        webView?.settings?.javaScriptEnabled = true
        webView?.webViewClient = object : WebViewClient() {
            override fun onPageFinished(view: WebView?, url: String?) {
                val gson = Gson()
                if (data !== null) {
                    webView?.loadUrl("javascript:setExcelData(${gson.toJson(data!!)}, ${data!![0].size})")
                    webView?.loadUrl("javascript:generateExcel()")
                }
                if (Build.VERSION.SDK_INT >= 19) {
                    if (!isPdf) {
                        createWebPrintJob(context, view)
                    } else {
                        val pdfPath = "${context.filesDir}${File.separator}testPDF.pdf"
                        printPDFFile(context, pdfPath, webView!!)
                    }
                } else {
                    Toast.makeText(context, "当前版本号低于 4.4", Toast.LENGTH_SHORT).show()
                }
            }
        }
        webView?.loadUrl(url)

    }

    @RequiresApi(Build.VERSION_CODES.KITKAT)
    private fun createWebPrintJob(context: Context, view: WebView?) {
        val printManager: PrintManager = context.getSystemService(Context.PRINT_SERVICE) as PrintManager

        val printDocumentAdapter = view?.createPrintDocumentAdapter()

        val jobName = "${R.string.app_name}Document"

        val printJob = printManager.print(jobName, printDocumentAdapter, PrintAttributes.Builder().build())
    }

    var dexCacheFile: File? = null
    // 获取需要打印的webview适配器
    var printAdapter: PrintDocumentAdapter? = null
    var ranges: Array<PageRange>? = null
    var descriptor: ParcelFileDescriptor? = null


    /**
     * a* @param webView
     */
    private fun printPDFFile(context: Context, pdfPath: String, webView: WebView) {
        var file = File(pdfPath)
        if (android.os.Build.VERSION.SDK_INT >= android.os.Build.VERSION_CODES.KITKAT) {
            /**
             * android 5.0之后，出于对动态注入字节码安全性德考虑，已经不允许随意指定字节码的保存路径了，需要放在应用自己的包名文件夹下。
             */
            //新的创建DexMaker缓存目录的方式，直接通过context获取路径
            dexCacheFile = context.getDir("dex", 0)
            if (!(dexCacheFile?.exists()!!)) {
                dexCacheFile?.mkdir()
            }

            try {
                //创建待写入的PDF文件，pdfFilePath为自行指定的PDF文件路径
                if (file.exists()) {
                    file.delete()
                }
                file.createNewFile()
                descriptor = ParcelFileDescriptor.open(file, ParcelFileDescriptor.MODE_READ_WRITE)

                // 设置打印参数
                val attributes = PrintAttributes.Builder()
                        .setMediaSize(PrintAttributes.MediaSize.ISO_A4)
                        .setResolution(PrintAttributes.Resolution("id", Context.PRINT_SERVICE, 300, 300))
                        .setColorMode(PrintAttributes.COLOR_MODE_COLOR)
                        .setMinMargins(PrintAttributes.Margins.NO_MARGINS)
                        .build()
                //打印所有界面
                ranges = arrayOf(PageRange.ALL_PAGES)

                printAdapter = webView.createPrintDocumentAdapter()
                // 开始打印
                printAdapter?.onStart()
                printAdapter?.onLayout(attributes, attributes, CancellationSignal(), getLayoutResultCallback(InvocationHandler { proxy, method, args ->
                    if (method.getName().equals("onLayoutFinished")) {
                        // 监听到内部调用了onLayoutFinished()方法，即打印成功
                        onLayoutSuccess(context)
                    } else {
                        // 监听到打印失败或者取消了打印

                    }
                    null
                }, dexCacheFile?.getAbsoluteFile()!!), Bundle())
            } catch (e: IOException) {
                e.printStackTrace()
            }

        }
    }

    /**
     * @throws IOException
     */
    @Throws(IOException::class)
    private fun onLayoutSuccess(context: Context) {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.KITKAT) {
            val callback = getWriteResultCallback(InvocationHandler { o, method, objects ->
                if (method.getName().equals("onWriteFinished")) {
                    Toast.makeText(context, "Success", Toast.LENGTH_SHORT).show()
                    // PDF文件写入本地完成，导出成功
                    Log.e("onLayoutSuccess", "onLayoutSuccess")
                    doPdfPrint(context, "${context.filesDir}${File.separator}testPDF.pdf")
                } else {
                    Toast.makeText(context, "导出失败", Toast.LENGTH_SHORT).show()
                }
                null
            }, dexCacheFile?.getAbsoluteFile()!!)
            //写入文件到本地
            printAdapter?.onWrite(ranges, descriptor, CancellationSignal(), callback)
        } else {
            Toast.makeText(context, "不支持4.4.以下", Toast.LENGTH_SHORT).show()

        }
    }

    @SuppressLint("NewApi")
    @Throws(IOException::class)
    fun getLayoutResultCallback(invocationHandler: InvocationHandler, dexCacheDir: File): PrintDocumentAdapter.LayoutResultCallback {
        return ProxyBuilder.forClass(PrintDocumentAdapter.LayoutResultCallback::class.java)
                .dexCache(dexCacheDir)
                .handler(invocationHandler)
                .build()
    }

    @SuppressLint("NewApi")
    @Throws(IOException::class)
    fun getWriteResultCallback(invocationHandler: InvocationHandler, dexCacheDir: File): PrintDocumentAdapter.WriteResultCallback {
        return ProxyBuilder.forClass(PrintDocumentAdapter.WriteResultCallback::class.java)
                .dexCache(dexCacheDir)
                .handler(invocationHandler)
                .build()
    }

    @RequiresApi(Build.VERSION_CODES.KITKAT)
    private fun doPdfPrint(context: Context, filePath: String) {
        var printManager = context.getSystemService(Context.PRINT_SERVICE) as PrintManager
        var myPrintAdapter = MyPrintPdfAdapter(filePath)
        printManager.print("jobName", myPrintAdapter, null)
    }

}