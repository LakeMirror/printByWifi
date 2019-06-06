package com.industry.cnwin.printerbywifi.testWebView

import android.app.ActivityManager
import android.content.pm.PackageManager
import android.os.Build
import android.os.Bundle
import android.support.v4.app.ActivityCompat
import android.support.v7.app.AppCompatActivity
import android.webkit.WebView
import android.webkit.WebViewClient
import com.google.gson.Gson
import com.industry.cnwin.printerbywifi.R
import java.util.jar.Manifest

class PreviewActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activty_preview)
        if (Build.VERSION.SDK_INT >= 23 && ActivityCompat.checkSelfPermission(this, android.Manifest.permission.INTERNET) == PackageManager.PERMISSION_DENIED) {
            requestPermissions(arrayOf(android.Manifest.permission.INTERNET), 1)
        } else {
            initView()
        }
    }

    val title = arrayOf("名字", "年纪", "班级", "性别", "身高", "体重", "成绩", "性格", "xx", "xx")
    private fun initView() {
        val preview = findViewById<WebView>(R.id.wv_preview)
        preview.webViewClient = WebViewClient()
        preview.settings.javaScriptEnabled = true
        preview.loadUrl("file:////android_asset/word.html")
        var data = Array<Array<String>>(10) { i -> Array(10) { i -> "" } }
        for (i in 1..9) {
            for (j in 0..9) {
                data[i][j] = (i * j).toString()
            }
        }
        data[0] = title

        preview.webViewClient = object : WebViewClient() {
            override fun onPageFinished(view: WebView?, url: String?) {
                super.onPageFinished(view, url)
                val gson = Gson()
                preview.loadUrl("javascript:setExcelData(${gson.toJson(data)}, ${data[0].size}, ${data.size})")
                preview.loadUrl("javascript:generateExcel()")
            }
        }
    }

    override fun onRequestPermissionsResult(requestCode: Int, permissions: Array<out String>, grantResults: IntArray) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        initView()
    }
}