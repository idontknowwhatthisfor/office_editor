package com.officeeditor

import android.content.Context
import android.net.Uri
import android.provider.OpenableColumns

/** The formats the app opens. [id] is what the web viewers switch on. */
enum class DocFormat(val id: String, val mime: String, val label: String) {
    PDF("pdf", "application/pdf", "PDF"),
    DOCX("docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "DOC"),
    XLSX("xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "XLS"),
    PPTX("pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", "PPT");

    companion object {
        val mimeTypes: Array<String> get() = entries.map { it.mime }.toTypedArray()

        /**
         * Providers often report a generic type (application/octet-stream), so the file extension
         * is the fallback.
         */
        fun detect(mime: String?, name: String?): DocFormat? =
            entries.firstOrNull { it.mime.equals(mime, ignoreCase = true) }
                ?: name?.substringAfterLast('.', "")?.lowercase()?.let { ext ->
                    entries.firstOrNull { it.id == ext }
                }
    }
}

data class DocumentInfo(
    val uri: Uri,
    val name: String,
    val size: Long,
    val format: DocFormat?,
)

/** Reads name, size and type for a content URI. Returns null if the file can't be reached. */
fun resolveDocument(context: Context, uri: Uri, mimeHint: String?): DocumentInfo? {
    val resolver = context.contentResolver
    var name: String? = null
    var size = -1L
    try {
        resolver.query(uri, null, null, null, null)?.use { cursor ->
            if (cursor.moveToFirst()) {
                val nameIdx = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                val sizeIdx = cursor.getColumnIndex(OpenableColumns.SIZE)
                if (nameIdx >= 0 && !cursor.isNull(nameIdx)) name = cursor.getString(nameIdx)
                if (sizeIdx >= 0 && !cursor.isNull(sizeIdx)) size = cursor.getLong(sizeIdx)
            }
        }
        // Some providers answer the query for files that are gone; opening is the real test.
        resolver.openAssetFileDescriptor(uri, "r")?.use { fd ->
            if (size < 0) size = fd.length
        } ?: return null
    } catch (e: Exception) {
        return null
    }
    val displayName = name ?: uri.lastPathSegment?.substringAfterLast('/')
    val format = DocFormat.detect(resolver.getType(uri), displayName)
        ?: DocFormat.detect(mimeHint, displayName)
    return DocumentInfo(uri, displayName ?: context.getString(R.string.untitled), size, format)
}
