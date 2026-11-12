package com.officeeditor

import android.annotation.SuppressLint
import android.content.Context
import android.content.pm.ApplicationInfo
import android.util.Log
import android.webkit.ConsoleMessage
import android.webkit.JavascriptInterface
import android.webkit.WebChromeClient
import android.webkit.WebResourceRequest
import android.webkit.WebResourceResponse
import android.webkit.WebView
import android.webkit.WebViewClient
import androidx.webkit.WebViewAssetLoader
import org.json.JSONObject

private const val TAG = "OfficeEditorWeb"
private const val VIEWER_URL = "https://${WebViewAssetLoader.DEFAULT_DOMAIN}/assets/web/index.html"
private const val DOC_PATH = "/doc/"

/**
 * Creates the WebView that hosts the bundled web viewers and hands it [doc].
 *
 * Everything is served from the app itself through [WebViewAssetLoader]: the viewer code from
 * `assets/web`, and the document's bytes as a stream at `/doc/current`, which the page fetches.
 * Streaming avoids base64-encoding whole files across the JS bridge. Any other URL is refused.
 */
@SuppressLint("SetJavaScriptEnabled")
fun createDocumentWebView(context: Context, doc: DocumentInfo): WebView {
    val format = requireNotNull(doc.format)
    val assetLoader = WebViewAssetLoader.Builder()
        .addPathHandler("/assets/", WebViewAssetLoader.AssetsPathHandler(context))
        .addPathHandler(DOC_PATH) {
            try {
                WebResourceResponse(format.mime, null, context.contentResolver.openInputStream(doc.uri))
            } catch (e: Exception) {
                Log.w(TAG, "Can't read ${doc.uri}", e)
                null
            }
        }
        .build()

    // Debug builds can be inspected from chrome://inspect on a connected computer.
    if (context.applicationInfo.flags and ApplicationInfo.FLAG_DEBUGGABLE != 0) {
        WebView.setWebContentsDebuggingEnabled(true)
    }

    return WebView(context).apply {
        settings.javaScriptEnabled = true
        settings.allowFileAccess = false
        settings.allowContentAccess = false
        // Pinch-zoom is the WebView's own; the page re-renders sharper once a zoom settles.
        settings.setSupportZoom(true)
        settings.builtInZoomControls = true
        settings.displayZoomControls = false

        webViewClient = object : WebViewClient() {
            override fun shouldInterceptRequest(view: WebView, request: WebResourceRequest): WebResourceResponse? =
                assetLoader.shouldInterceptRequest(request.url)
                    ?: WebResourceResponse("text/plain", "utf-8", 404, "Not found", emptyMap(), null)

            // The viewer is a single page; nothing in a document may navigate away from it.
            override fun shouldOverrideUrlLoading(view: WebView, request: WebResourceRequest): Boolean = true
        }
        webChromeClient = object : WebChromeClient() {
            override fun onConsoleMessage(message: ConsoleMessage): Boolean {
                val line = "${message.message()} (${message.sourceId()}:${message.lineNumber()})"
                when (message.messageLevel()) {
                    ConsoleMessage.MessageLevel.ERROR -> Log.e(TAG, line)
                    ConsoleMessage.MessageLevel.WARNING -> Log.w(TAG, line)
                    else -> Log.d(TAG, line)
                }
                return true
            }
        }
        addJavascriptInterface(Bridge(doc, format), "OfficeBridge")
        loadUrl(VIEWER_URL)
    }
}

/** What the web viewers can ask the native shell. Called on a WebView thread, not the UI thread. */
private class Bridge(private val doc: DocumentInfo, private val format: DocFormat) {
    /** The open document as JSON: `{name, format, mime, size, url}`; fetch `url` for the bytes. */
    @JavascriptInterface
    fun getDocument(): String = JSONObject()
        .put("name", doc.name)
        .put("format", format.id)
        .put("mime", format.mime)
        .put("size", doc.size)
        .put("url", DOC_PATH + "current")
        .toString()
}
