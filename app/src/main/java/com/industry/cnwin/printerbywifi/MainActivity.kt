package com.industry.cnwin.printerbywifi

import android.annotation.SuppressLint
import android.content.Context
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.os.Build
import android.support.v7.app.AppCompatActivity
import android.os.Bundle
import android.print.PrintManager
import android.renderscript.ScriptGroup
import android.support.v4.print.PrintHelper
import android.util.Log
import android.view.View
import android.widget.Button
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.InputStream
import java.util.jar.Manifest

class MainActivity : AppCompatActivity(), View.OnClickListener {
    override fun onClick(v: View?) {
        val data: Array<Array<String>> = Array(10){ i -> Array(10) {j -> ""} }
        for (i in 1..9) {
            for (j in 0..9) {
                data[i][j] = (i * j).toString()
            }
        }
        data[0] =  arrayOf("姓名", "年龄", "性别", "班级", "楼层", "成绩", "性格", "为人", "", "")

        when(v?.id){
            R.id.bt_print_bitmap -> {
                val bitmap = BitmapFactory.decodeResource(resources, R.drawable.app_image_small)
                PrintService.instance.printBitmap(this@MainActivity, bitmap)
            }
            R.id.bt_print_HTML -> {
                val url = "file:////android_asset/word.html"

                PrintService.instance.printExcelByWebView(this@MainActivity, url, data)
            }
            R.id.bt_print_PDF -> {
                val url = "file:////android_asset/word.html"
                PrintService.instance.printPDFByWebView(this@MainActivity, url, data)
            }
            R.id.bt_print_word -> {
                var file: File = File("${filesDir}${File.separator}test.png")
                var os = FileOutputStream(file)
                var inputStream = assets.open("app_image_small.png");
                var bts = ByteArray(1024)
                var len = inputStream.read(bts)
                while (len != -1) {
                    os.write(bts, 0, len)
                    os.flush()
                    len = inputStream.read(bts)
                }
                var path = WordGenerate.newInstance(this).newWordFile("test").insertTitle("这是一个测试的 Word")
                        .insertWord(ParagraphAlignment.LEFT, 16, false, "测试测试，好开心好开心xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
//                        .insertPicture(FileInputStream(file))
                        .insertExcel(data)
                        .generate()
                PrintService.instance.printWord(this@MainActivity, path)

            }
        }

    }


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        if (Build.VERSION.SDK_INT > 22)
            requestPermissions(arrayOf(android.Manifest.permission.READ_EXTERNAL_STORAGE, android.Manifest.permission.WRITE_APN_SETTINGS), 0)
        intiView()
    }

    private fun intiView() {
        val btPrintBitmap = findViewById<Button>(R.id.bt_print_bitmap)
        btPrintBitmap.setOnClickListener(this)
        val btPrintHTML = findViewById<Button>(R.id.bt_print_HTML)
        btPrintHTML.setOnClickListener(this)
        val btPrintPDF = findViewById<Button>(R.id.bt_print_PDF)
        btPrintPDF.setOnClickListener(this)
        val btPrintWord = findViewById<Button>(R.id.bt_print_word)
        btPrintWord.setOnClickListener(this)
    }
}
