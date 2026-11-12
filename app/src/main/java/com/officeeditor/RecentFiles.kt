package com.officeeditor

import android.content.Context
import android.content.Intent
import android.net.Uri
import java.io.File
import java.security.MessageDigest
import org.json.JSONArray
import org.json.JSONObject

data class RecentFile(val uri: Uri, val name: String, val format: DocFormat, val openedAt: Long)

/** The home screen's recent-files list, newest first, kept in SharedPreferences. */
class RecentFiles(context: Context) {
    private val context = context.applicationContext
    private val prefs = this.context.getSharedPreferences("recents", Context.MODE_PRIVATE)

    fun list(): List<RecentFile> {
        val json = prefs.getString(KEY, null) ?: return emptyList()
        return try {
            val array = JSONArray(json)
            (0 until array.length()).mapNotNull { i ->
                val o = array.getJSONObject(i)
                val format = DocFormat.entries.firstOrNull { it.id == o.getString("format") }
                    ?: return@mapNotNull null
                RecentFile(Uri.parse(o.getString("uri")), o.getString("name"), format, o.getLong("openedAt"))
            }
        } catch (e: Exception) {
            emptyList()
        }
    }

    /**
     * Records [doc] as just opened. Call off the main thread: it may copy the file.
     *
     * A file handed to us by another app ("Open with", share) comes with access that ends when
     * the viewer closes, so it couldn't be reopened from the list later. Unless the access can be
     * made lasting (files from the system picker), the list points at a private copy instead.
     */
    fun add(doc: DocumentInfo) {
        val format = doc.format ?: return
        val uri = reopenableUri(doc) ?: return
        val entry = RecentFile(uri, doc.name, format, System.currentTimeMillis())
        save((listOf(entry) + list().filter { it.uri != uri }).take(MAX))
    }

    private fun reopenableUri(doc: DocumentInfo): Uri? {
        if (isOwnCopy(context, doc.uri)) return doc.uri
        val resolver = context.contentResolver
        try {
            resolver.takePersistableUriPermission(doc.uri, Intent.FLAG_GRANT_READ_URI_PERMISSION)
        } catch (e: SecurityException) {
            // Not offered by the sender; fall through to the copy.
        }
        if (resolver.persistedUriPermissions.any { it.uri == doc.uri && it.isReadPermission }) return doc.uri
        if (doc.size > MAX_COPY_BYTES) return null

        return try {
            // One folder per source URI, so reopening the same attachment replaces its copy.
            val dir = File(copiesDir(context), sha1(doc.uri.toString())).apply {
                deleteRecursively()
                mkdirs()
            }
            val file = File(dir, doc.name.replace('/', '_'))
            resolver.openInputStream(doc.uri)!!.use { input ->
                file.outputStream().use { input.copyTo(it) }
            }
            Uri.fromFile(file)
        } catch (e: Exception) {
            null
        }
    }

    fun remove(uri: Uri) = save(list().filter { it.uri != uri })

    private fun save(files: List<RecentFile>) {
        val array = JSONArray()
        for (f in files) {
            array.put(
                JSONObject()
                    .put("uri", f.uri.toString())
                    .put("name", f.name)
                    .put("format", f.format.id)
                    .put("openedAt", f.openedAt)
            )
        }
        prefs.edit().putString(KEY, array.toString()).apply()

        // Drop private copies that are no longer in the list.
        val kept = files.mapNotNull { it.uri.path?.let(::File)?.parentFile?.canonicalPath }.toSet()
        copiesDir(context).listFiles()?.forEach { if (it.canonicalPath !in kept) it.deleteRecursively() }
    }

    private fun sha1(text: String): String =
        MessageDigest.getInstance("SHA-1").digest(text.toByteArray()).joinToString("") { "%02x".format(it) }

    companion object {
        private const val KEY = "files"
        private const val MAX = 30
        private const val MAX_COPY_BYTES = 50L * 1024 * 1024

        private fun copiesDir(context: Context) = File(context.filesDir, "recent")

        /** Whether [uri] is one of the private copies kept for the recent-files list. */
        fun isOwnCopy(context: Context, uri: Uri): Boolean {
            if (uri.scheme != "file") return false
            val path = uri.path ?: return false
            return File(path).canonicalPath.startsWith(copiesDir(context).canonicalPath + File.separator)
        }
    }
}
