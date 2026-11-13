package com.officeeditor

import android.content.Context
import android.os.Bundle
import android.os.CancellationSignal
import android.os.ParcelFileDescriptor
import android.print.PageRange
import android.print.PrintAttributes
import android.print.PrintDocumentAdapter
import android.print.PrintManager
import android.webkit.WebView
import kotlin.math.abs

// "Export as PDF": a WebView holding the document's print layout (web/src/print.ts) is printed
// through Android's print framework. Its "Save as PDF" printer writes the file wherever the user
// chooses; any installed printer works too. Nothing leaves the device unless the user picks a
// network printer.

/**
 * Opens the system print dialog for [view], with the paper set to the document's page size
 * ([widthPx] × [heightPx] CSS px, 96 per inch) and no margins (the layout has its own).
 * [onFinish] runs once the dialog is done with the WebView.
 */
fun printWebView(context: Context, view: WebView, jobName: String, widthPx: Double, heightPx: Double, onFinish: () -> Unit) {
    val attributes = PrintAttributes.Builder()
        .setMediaSize(mediaSizeFor(widthPx, heightPx))
        .setMinMargins(PrintAttributes.Margins.NO_MARGINS)
        .setColorMode(PrintAttributes.COLOR_MODE_COLOR)
        .build()
    val adapter = FinishingAdapter(view.createPrintDocumentAdapter(jobName), onFinish)
    context.getSystemService(PrintManager::class.java).print(jobName, adapter, attributes)
}

/**
 * The standard paper size matching the page. Android's PDF printer ignores custom sizes, so
 * anything else gets A4 the same way round (slides are already laid out on landscape A4).
 */
private fun mediaSizeFor(widthPx: Double, heightPx: Double): PrintAttributes.MediaSize {
    val w = (widthPx * 1000 / 96).toInt()
    val h = (heightPx * 1000 / 96).toInt()
    val standard = listOf(
        PrintAttributes.MediaSize.ISO_A4, PrintAttributes.MediaSize.NA_LETTER, PrintAttributes.MediaSize.NA_LEGAL,
        PrintAttributes.MediaSize.ISO_A5, PrintAttributes.MediaSize.ISO_A3, PrintAttributes.MediaSize.ISO_B5,
    ).flatMap { listOf(it.asPortrait(), it.asLandscape()) }
    fun near(a: Int, b: Int) = abs(a - b) <= b * 0.01
    return standard.firstOrNull { near(w, it.widthMils) && near(h, it.heightMils) }
        ?: if (w > h) PrintAttributes.MediaSize.ISO_A4.asLandscape() else PrintAttributes.MediaSize.ISO_A4
}

/** Passes everything to the WebView's adapter and reports when printing is over. */
private class FinishingAdapter(private val inner: PrintDocumentAdapter, private val onDone: () -> Unit) : PrintDocumentAdapter() {
    override fun onStart() = inner.onStart()

    override fun onLayout(
        oldAttributes: PrintAttributes?,
        newAttributes: PrintAttributes,
        cancellationSignal: CancellationSignal?,
        callback: LayoutResultCallback,
        extras: Bundle?,
    ) = inner.onLayout(oldAttributes, newAttributes, cancellationSignal, callback, extras)

    override fun onWrite(
        pages: Array<out PageRange>,
        destination: ParcelFileDescriptor,
        cancellationSignal: CancellationSignal?,
        callback: WriteResultCallback,
    ) = inner.onWrite(pages, destination, cancellationSignal, callback)

    override fun onFinish() {
        inner.onFinish()
        onDone()
    }
}
