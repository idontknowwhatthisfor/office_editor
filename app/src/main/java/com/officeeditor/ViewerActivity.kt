package com.officeeditor

import android.content.ClipData
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.webkit.WebView
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.BackHandler
import androidx.activity.compose.setContent
import androidx.activity.result.ActivityResultLauncher
import androidx.activity.result.contract.ActivityResultContracts
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.WindowInsets
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.ime
import androidx.compose.foundation.layout.navigationBars
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.union
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Share
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.size
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.TextButton
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.setValue
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.material3.TopAppBar
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.res.stringResource
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.viewinterop.AndroidView
import androidx.core.content.FileProvider
import java.io.File
import kotlin.concurrent.thread

/**
 * Shows one document. Reached from the home screen, from "Open with" (ACTION_VIEW) and from the
 * share sheet (ACTION_SEND).
 */
class ViewerActivity : ComponentActivity(), EditorHost {
    private var webView: WebView? = null
    private var doc: DocumentInfo? = null
    private var unsaved by mutableStateOf(false)
    private var saving by mutableStateOf(false)
    private var writable by mutableStateOf(false)
    private var dialog by mutableStateOf(Dialog.NONE)
    /** Leave the screen once the save in progress succeeds (the user chose Save on leaving). */
    private var closeAfterSave = false
    private var pendingCopy: File? = null
    private lateinit var createCopy: ActivityResultLauncher<String>
    override val saves by lazy { SaveReceiver(this) }

    private enum class Dialog { NONE, SAVE, UNSAVED }

    @OptIn(ExperimentalMaterial3Api::class)
    override fun onCreate(savedInstanceState: Bundle?) {
        enableEdgeToEdge()
        super.onCreate(savedInstanceState)

        val uri = documentUri(intent)
        val doc = uri?.let { resolveDocument(this, it, intent.type) }
        val error = when {
            doc == null -> getString(R.string.file_unavailable)
            doc.format == null -> getString(R.string.unsupported_format)
            else -> null
        }
        this.doc = doc
        createCopy = registerForActivityResult(
            ActivityResultContracts.CreateDocument(doc?.format?.mime ?: "application/octet-stream"),
        ) { target -> copyChosen(target) }
        val recents = RecentFiles(this)
        if (doc != null && error == null) {
            thread {
                recents.add(doc)
                val canWrite = canOverwrite(this, doc.uri)
                runOnUiThread { writable = canWrite }
            }
            webView = createDocumentWebView(this, doc, this)
        } else if (uri != null && doc == null) {
            recents.remove(uri)
        }

        setContent {
            OfficeEditorTheme {
                Scaffold(
                    topBar = {
                        TopAppBar(
                            title = {
                                Text(
                                    doc?.name ?: stringResource(R.string.app_name),
                                    maxLines = 1,
                                    overflow = TextOverflow.Ellipsis,
                                )
                            },
                            navigationIcon = {
                                IconButton(onClick = { finish() }) {
                                    Icon(Icons.AutoMirrored.Filled.ArrowBack, stringResource(R.string.back))
                                }
                            },
                            actions = {
                                if (saving) {
                                    CircularProgressIndicator(Modifier.padding(12.dp).size(24.dp), strokeWidth = 2.dp)
                                } else if (unsaved) {
                                    TextButton(onClick = { dialog = Dialog.SAVE }) { Text(stringResource(R.string.save)) }
                                }
                                if (doc != null && error == null) {
                                    IconButton(onClick = { share(doc) }) {
                                        Icon(Icons.Filled.Share, stringResource(R.string.share))
                                    }
                                }
                            },
                        )
                    },
                    // The web page lays out above the navigation bar and the keyboard.
                    // (Dialogs are drawn after the content, below.)
                    contentWindowInsets = WindowInsets.navigationBars.union(WindowInsets.ime),
                ) { padding ->
                    val view = webView
                    if (view != null) {
                        AndroidView(factory = { view }, modifier = Modifier.fillMaxSize().padding(padding))
                    } else {
                        Column(
                            modifier = Modifier.fillMaxSize().padding(padding).padding(32.dp),
                            verticalArrangement = Arrangement.Center,
                            horizontalAlignment = Alignment.CenterHorizontally,
                        ) {
                            Text(stringResource(R.string.cant_open_title), style = MaterialTheme.typography.titleMedium)
                            Spacer(Modifier.height(8.dp))
                            Text(
                                error.orEmpty(),
                                style = MaterialTheme.typography.bodyMedium,
                                color = MaterialTheme.colorScheme.onSurfaceVariant,
                                textAlign = TextAlign.Center,
                            )
                        }
                    }
                }
                BackHandler(enabled = unsaved && !saving) { dialog = Dialog.UNSAVED }
                SaveDialogs()
            }
        }
    }

    @androidx.compose.runtime.Composable
    private fun SaveDialogs() {
        when (dialog) {
            Dialog.NONE -> Unit
            Dialog.SAVE -> AlertDialog(
                onDismissRequest = {
                    dialog = Dialog.NONE
                    closeAfterSave = false
                },
                title = { Text(stringResource(R.string.save_title)) },
                text = { Text(stringResource(if (writable) R.string.save_body else R.string.save_body_copy_only)) },
                confirmButton = {
                    Row {
                        TextButton(onClick = {
                            dialog = Dialog.NONE
                            closeAfterSave = false
                        }) { Text(stringResource(R.string.cancel)) }
                        TextButton(onClick = { startSave("copy") }) { Text(stringResource(R.string.save_copy)) }
                        if (writable) {
                            TextButton(onClick = { startSave("overwrite") }) { Text(stringResource(R.string.save_overwrite)) }
                        }
                    }
                },
            )
            Dialog.UNSAVED -> AlertDialog(
                onDismissRequest = { dialog = Dialog.NONE },
                title = { Text(stringResource(R.string.unsaved_title)) },
                text = { Text(stringResource(R.string.unsaved_body)) },
                confirmButton = {
                    Row {
                        TextButton(onClick = { dialog = Dialog.NONE }) { Text(stringResource(R.string.cancel)) }
                        TextButton(onClick = {
                            dialog = Dialog.NONE
                            finish()
                        }) { Text(stringResource(R.string.discard)) }
                        TextButton(onClick = {
                            closeAfterSave = true
                            dialog = Dialog.SAVE
                        }) { Text(stringResource(R.string.save)) }
                    }
                },
            )
        }
    }

    // ---- saving: the web editor serializes and re-checks the file, then streams it here ----

    override fun setDirty(dirty: Boolean) {
        runOnUiThread { unsaved = dirty }
    }

    private fun startSave(mode: String) {
        dialog = Dialog.NONE
        saving = true
        webView?.evaluateJavascript("window.officeSave ? window.officeSave('$mode') : 'none'") { result ->
            if (result == "\"none\"") saveFailed()
        }
    }

    override fun finishSave(token: Int, mode: String) {
        val doc = doc ?: return
        val file = try {
            saves.finish(token, doc.format!!)
        } catch (e: Exception) {
            null
        }
        if (file == null) {
            runOnUiThread { saveFailed() }
            return
        }
        if (mode == "copy") {
            runOnUiThread {
                pendingCopy = file
                createCopy.launch(copyName(doc.name))
            }
            return
        }
        val ok = overwrite(this, doc.uri, file)
        runOnUiThread {
            if (!ok) return@runOnUiThread saveFailed()
            saving = false
            unsaved = false
            Toast.makeText(this, R.string.saved, Toast.LENGTH_SHORT).show()
            notifyWeb(true, mode)
            if (closeAfterSave) finish()
        }
    }

    /** The user picked where the copy goes (or cancelled: [target] is null). */
    private fun copyChosen(target: Uri?) {
        val file = pendingCopy
        pendingCopy = null
        if (target == null || file == null) {
            saving = false
            closeAfterSave = false
            notifyWeb(false, "copy", cancelled = true)
            return
        }
        thread {
            val ok = writeCopy(this, target, file)
            val copy = if (ok) resolveDocument(this, target, doc?.format?.mime) else null
            if (copy != null) RecentFiles(this).add(copy)
            runOnUiThread {
                if (copy == null) return@runOnUiThread saveFailed()
                saving = false
                unsaved = false
                Toast.makeText(this, R.string.saved_copy, Toast.LENGTH_SHORT).show()
                if (!closeAfterSave) {
                    // Carry on working in the copy, so the next save goes there.
                    startActivity(
                        Intent(this, ViewerActivity::class.java).setAction(Intent.ACTION_VIEW)
                            .setDataAndType(target, copy.format?.mime)
                            .addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION),
                    )
                }
                finish()
            }
        }
    }

    private fun saveFailed() {
        saving = false
        closeAfterSave = false
        Toast.makeText(this, R.string.save_failed, Toast.LENGTH_LONG).show()
        notifyWeb(false, "")
    }

    private fun notifyWeb(ok: Boolean, mode: String, cancelled: Boolean = false) {
        webView?.evaluateJavascript(
            "window.officeSaveResult && window.officeSaveResult({ok:$ok,mode:'$mode',cancelled:$cancelled})", null,
        )
    }

    override fun onDestroy() {
        webView?.destroy()
        webView = null
        super.onDestroy()
    }

    private fun documentUri(intent: Intent): Uri? = when {
        // This activity is exported, so only accept paths that really are our own copies.
        intent.hasExtra(EXTRA_OWN_COPY) ->
            Uri.fromFile(File(intent.getStringExtra(EXTRA_OWN_COPY).orEmpty()))
                .takeIf { RecentFiles.isOwnCopy(this, it) }
        intent.action == Intent.ACTION_SEND ->
            if (Build.VERSION.SDK_INT >= 33) {
                intent.getParcelableExtra(Intent.EXTRA_STREAM, Uri::class.java)
            } else {
                @Suppress("DEPRECATION")
                intent.getParcelableExtra(Intent.EXTRA_STREAM)
            }
        else -> intent.data
    }

    companion object {
        /** Path of a private copy kept by [RecentFiles], used instead of intent data. */
        const val EXTRA_OWN_COPY = "com.officeeditor.OWN_COPY"
    }

    /**
     * Shares a copy from our cache rather than the original URI: access granted to us by another
     * app (e.g. a chat attachment) can't be passed on to a third app.
     */
    private fun share(doc: DocumentInfo) {
        val format = doc.format ?: return
        thread {
            val shared = try {
                val dir = File(cacheDir, "shared").apply {
                    deleteRecursively()
                    mkdirs()
                }
                val file = File(dir, doc.name.replace('/', '_'))
                contentResolver.openInputStream(doc.uri)!!.use { input ->
                    file.outputStream().use { input.copyTo(it) }
                }
                FileProvider.getUriForFile(this, "$packageName.share", file)
            } catch (e: Exception) {
                null
            }
            runOnUiThread {
                if (shared == null) {
                    Toast.makeText(this, R.string.share_failed, Toast.LENGTH_SHORT).show()
                    return@runOnUiThread
                }
                val send = Intent(Intent.ACTION_SEND)
                    .setType(format.mime)
                    .putExtra(Intent.EXTRA_STREAM, shared)
                    .addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                // The chooser passes the read grant on only for URIs in the clip data.
                send.clipData = ClipData.newRawUri(null, shared)
                startActivity(Intent.createChooser(send, null))
            }
        }
    }
}
