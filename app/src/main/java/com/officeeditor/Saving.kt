package com.officeeditor

import android.content.Context
import android.net.Uri
import android.util.Base64
import android.util.Log
import java.io.File
import java.io.OutputStream

private const val TAG = "OfficeEditorSave"

/**
 * Receives an edited file from the web editor in base64 chunks (the JS bridge passes strings
 * only) and writes it to a temporary file. Nothing touches the original until the whole new file
 * is on disk and looks like the right format; the editor has already re-parsed it before sending.
 */
class SaveReceiver(context: Context) {
    private val dir = File(context.cacheDir, "saving")
    private var token = 0
    private var file: File? = null
    private var out: OutputStream? = null

    @Synchronized
    fun begin(): Int {
        out?.close()
        dir.deleteRecursively()
        dir.mkdirs()
        file = File.createTempFile("edit", ".tmp", dir)
        out = file!!.outputStream().buffered()
        return ++token
    }

    @Synchronized
    fun write(token: Int, base64: String) {
        check(token == this.token) { "stale save" }
        out!!.write(Base64.decode(base64, Base64.NO_WRAP))
    }

    /** The finished file, or null if it isn't a well-formed [format] file. */
    @Synchronized
    fun finish(token: Int, format: DocFormat): File? {
        check(token == this.token) { "stale save" }
        out!!.close()
        out = null
        val result = file!!
        val head = ByteArray(5)
        val read = result.inputStream().use { it.read(head) }
        val ok = when (format) {
            DocFormat.PDF -> read == 5 && String(head, Charsets.ISO_8859_1) == "%PDF-"
            else -> read >= 4 && head[0] == 'P'.code.toByte() && head[1] == 'K'.code.toByte() &&
                head[2] == 3.toByte() && head[3] == 4.toByte()
        }
        if (!ok) Log.w(TAG, "edited file is not a valid ${format.id}")
        return result.takeIf { ok && it.length() > 0 }
    }
}

/** Whether we may write to [uri] (and it is the user's file, not our private copy of one). */
fun canOverwrite(context: Context, uri: Uri): Boolean {
    if (RecentFiles.isOwnCopy(context, uri)) return false
    return try {
        context.contentResolver.openFileDescriptor(uri, "rw")?.close() != null
    } catch (e: Exception) {
        false
    }
}

/**
 * Replaces the contents of [uri] with [file]. The original is copied aside first; if writing
 * fails part-way, it is put back.
 */
fun overwrite(context: Context, uri: Uri, file: File): Boolean {
    val resolver = context.contentResolver
    val backup = File(context.cacheDir, "backup").run {
        deleteRecursively()
        mkdirs()
        File(this, "original")
    }
    return try {
        resolver.openInputStream(uri)!!.use { input -> backup.outputStream().use { input.copyTo(it) } }
        try {
            resolver.openOutputStream(uri, "wt")!!.use { out -> file.inputStream().use { it.copyTo(out) } }
            val written = resolver.openFileDescriptor(uri, "r")?.use { it.statSize } ?: -1
            check(written < 0 || written == file.length()) { "wrote $written of ${file.length()} bytes" }
            true
        } catch (e: Exception) {
            Log.e(TAG, "overwrite failed; restoring the original", e)
            try {
                resolver.openOutputStream(uri, "wt")!!.use { out -> backup.inputStream().use { it.copyTo(out) } }
            } catch (restore: Exception) {
                Log.e(TAG, "restore failed; the original is kept at ${backup.path}", restore)
                return false
            }
            false
        }
    } catch (e: Exception) {
        Log.e(TAG, "overwrite failed before writing", e)
        false
    }
}

fun writeCopy(context: Context, uri: Uri, file: File): Boolean = try {
    context.contentResolver.openOutputStream(uri, "wt")!!.use { out -> file.inputStream().use { it.copyTo(out) } }
    true
} catch (e: Exception) {
    Log.e(TAG, "save as copy failed", e)
    false
}

/** "Report.pdf" → "Report (edited).pdf". */
fun copyName(name: String): String {
    val dot = name.lastIndexOf('.')
    return if (dot > 0) "${name.substring(0, dot)} (edited)${name.substring(dot)}" else "$name (edited)"
}
