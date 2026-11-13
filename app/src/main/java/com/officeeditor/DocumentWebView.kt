package com.officeeditor

import android.annotation.SuppressLint
import android.content.Context
import android.content.pm.ApplicationInfo
import android.net.Uri
import android.os.ParcelFileDescriptor
import android.system.Os
import android.system.OsConstants
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
import java.io.InputStream

private const val TAG = "OfficeEditorWeb"
private const val VIEWER_URL = "https://${WebViewAssetLoader.DEFAULT_DOMAIN}/assets/web/index.html"
private const val DOC_PATH = "/doc/"

/**
 * Creates the WebView that hosts the bundled web viewers and hands it [doc].
 *
 * Everything is served from the app itself through [WebViewAssetLoader]: the viewer code from
 * `assets/web`, and the document's bytes at `/doc/current`, which the page fetches (whole, or in
 * ranges). That avoids base64-encoding whole files across the JS bridge. Any other URL is refused.
 */
@SuppressLint("SetJavaScriptEnabled")
fun createDocumentWebView(context: Context, doc: DocumentInfo, host: EditorHost, startEditing: Boolean = false): WebView {
    val format = requireNotNull(doc.format)
    val assetLoader = WebViewAssetLoader.Builder()
        .addPathHandler("/assets/", WebViewAssetLoader.AssetsPathHandler(context))
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
            override fun shouldInterceptRequest(view: WebView, request: WebResourceRequest): WebResourceResponse? {
                val url = request.url
                val response =
                    if (url.host == WebViewAssetLoader.DEFAULT_DOMAIN && url.path?.startsWith(DOC_PATH) == true) {
                        serveDocument(context, doc.uri, format.mime, request)
                    } else {
                        assetLoader.shouldInterceptRequest(url)
                    }
                return response ?: WebResourceResponse("text/plain", "utf-8", 404, "Not found", emptyMap(), null)
            }

            // The viewer is a single page; nothing in a document may navigate away from it. Only
            // reloading the viewer itself is allowed (an editor does that to show the saved file).
            override fun shouldOverrideUrlLoading(view: WebView, request: WebResourceRequest): Boolean =
                request.url.toString().substringBefore('#') != VIEWER_URL
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
        addJavascriptInterface(Bridge(doc, format, host, startEditing), "OfficeBridge")
        loadUrl(VIEWER_URL)
    }
}

/**
 * Answers a request for the document's bytes. Honours HTTP `Range` requests when the file is
 * seekable, so a viewer can read a large file piece by piece instead of holding all of it in
 * memory (the PDF viewer does). Providers that only offer a pipe get the whole stream.
 */
private fun serveDocument(context: Context, uri: Uri, mime: String, request: WebResourceRequest): WebResourceResponse? {
    try {
        val resolver = context.contentResolver
        val descriptor = resolver.openFileDescriptor(uri, "r") ?: return null
        val size = descriptor.statSize
        val seekable = size >= 0 && try {
            Os.lseek(descriptor.fileDescriptor, 0, OsConstants.SEEK_CUR)
            true
        } catch (e: Exception) {
            false
        }
        if (!seekable) {
            descriptor.close()
            return WebResourceResponse(mime, null, resolver.openInputStream(uri))
        }

        val stream = ParcelFileDescriptor.AutoCloseInputStream(descriptor)
        val headers = mutableMapOf("Accept-Ranges" to "bytes")
        val range = request.requestHeaders.entries.firstOrNull { it.key.equals("Range", ignoreCase = true) }?.value
        val match = range?.let { Regex("bytes=(\\d+)-(\\d*)").matchEntire(it.trim()) }
        // No Content-Length header: the WebView adds its own.
        if (match == null) {
            return WebResourceResponse(mime, null, 200, "OK", headers, stream)
        }
        val start = match.groupValues[1].toLong()
        val end = minOf(match.groupValues[2].toLongOrNull() ?: Long.MAX_VALUE, size - 1)
        if (start > end) {
            stream.close()
            headers["Content-Range"] = "bytes */$size"
            return WebResourceResponse(mime, null, 416, "Range Not Satisfiable", headers, null)
        }
        headers["Content-Range"] = "bytes $start-$end/$size"
        return WebResourceResponse(mime, null, 206, "Partial Content", headers, EndingAt(stream, end + 1))
    } catch (e: Exception) {
        Log.w(TAG, "Can't read $uri", e)
        return null
    }
}

/**
 * [source] cut off at byte offset [end] of the file, for answering a Range request.
 *
 * For a request with a Range header the WebView itself skips to the start of the range (so the
 * stream must begin at byte 0 of the file, and a stream of just the slice fails), but it then
 * reads to the end of the stream instead of the end of the range. So the stream ends there.
 */
private class EndingAt(private val source: InputStream, private val end: Long) : InputStream() {
    private var position = 0L

    override fun skip(count: Long): Long = source.skip(count).also { position += it }

    override fun read(): Int {
        if (position >= end) return -1
        return source.read().also { if (it >= 0) position++ }
    }

    override fun read(buffer: ByteArray, offset: Int, length: Int): Int {
        if (position >= end) return -1
        val count = source.read(buffer, offset, minOf(length.toLong(), end - position).toInt())
        if (count > 0) position += count
        return count
    }

    override fun available(): Int = minOf(end - position, Int.MAX_VALUE.toLong()).toInt()

    override fun close() = source.close()
}

/** What the viewer screen does for the web editors. Called on a WebView thread. */
interface EditorHost {
    /** The document has (or no longer has) unsaved changes. */
    fun setDirty(dirty: Boolean)
    val saves: SaveReceiver
    /** All chunks of an edited file have arrived; [mode] is "overwrite" or "copy". */
    fun finishSave(token: Int, mode: String)
}

/** What the web viewers can ask the native shell. Called on a WebView thread, not the UI thread. */
private class Bridge(
    private val doc: DocumentInfo,
    private val format: DocFormat,
    private val host: EditorHost,
    private val startEditing: Boolean,
) {
    /**
     * The open document as JSON: `{name, format, mime, size, url, edit}`; fetch `url` for the
     * bytes. `edit`: open in edit mode.
     */
    @JavascriptInterface
    fun getDocument(): String = JSONObject()
        .put("name", doc.name)
        .put("format", format.id)
        .put("mime", format.mime)
        .put("size", doc.size)
        .put("url", DOC_PATH + "current")
        .put("edit", startEditing)
        .toString()

    @JavascriptInterface
    fun setDirty(dirty: Boolean) = host.setDirty(dirty)

    /** Starts receiving an edited file; returns a token for [writeChunk] and [finishSave]. */
    @JavascriptInterface
    fun beginSave(): Int = host.saves.begin()

    @JavascriptInterface
    fun writeChunk(token: Int, base64: String) = host.saves.write(token, base64)

    @JavascriptInterface
    fun finishSave(token: Int, mode: String) = host.finishSave(token, mode)
}
