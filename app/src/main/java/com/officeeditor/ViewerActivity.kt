package com.officeeditor

import android.content.ClipData
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.webkit.WebView
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
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
import androidx.compose.material3.ExperimentalMaterial3Api
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
class ViewerActivity : ComponentActivity() {
    private var webView: WebView? = null

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
        val recents = RecentFiles(this)
        if (doc != null && error == null) {
            thread { recents.add(doc) }
            webView = createDocumentWebView(this, doc)
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
                                if (doc != null && error == null) {
                                    IconButton(onClick = { share(doc) }) {
                                        Icon(Icons.Filled.Share, stringResource(R.string.share))
                                    }
                                }
                            },
                        )
                    },
                    // The web page lays out above the navigation bar and the keyboard.
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
            }
        }
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
