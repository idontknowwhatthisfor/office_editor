package com.officeeditor

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.text.format.DateUtils
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.PaddingValues
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Add
import androidx.compose.material.icons.filled.Close
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.DropdownMenu
import androidx.compose.material3.DropdownMenuItem
import androidx.compose.material3.ExtendedFloatingActionButton
import androidx.compose.material3.FloatingActionButton
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.material3.TopAppBar
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.res.stringResource
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import kotlin.concurrent.thread

/** Home: recent files, "Open file", and new blank documents and spreadsheets. */
class MainActivity : ComponentActivity() {
    private lateinit var recents: RecentFiles
    private var files by mutableStateOf(emptyList<RecentFile>())

    override fun onCreate(savedInstanceState: Bundle?) {
        enableEdgeToEdge()
        super.onCreate(savedInstanceState)
        recents = RecentFiles(this)
        setContent {
            OfficeEditorTheme {
                val picker = rememberLauncherForActivityResult(ActivityResultContracts.OpenDocument()) { uri ->
                    if (uri != null) open(uri, null)
                }
                // "New": the user picks where the file goes, then it starts as a copy of a blank one.
                val newDocument = rememberLauncherForActivityResult(ActivityResultContracts.CreateDocument(DocFormat.DOCX.mime)) {
                    if (it != null) create(it, DocFormat.DOCX)
                }
                val newSpreadsheet = rememberLauncherForActivityResult(ActivityResultContracts.CreateDocument(DocFormat.XLSX.mime)) {
                    if (it != null) create(it, DocFormat.XLSX)
                }
                HomeScreen(
                    files = files,
                    onOpenPicker = { picker.launch(DocFormat.mimeTypes) },
                    onNewDocument = { newDocument.launch(getString(R.string.new_document_name) + ".docx") },
                    onNewSpreadsheet = { newSpreadsheet.launch(getString(R.string.new_spreadsheet_name) + ".xlsx") },
                    onOpen = { open(it.uri, it.format.mime) },
                    onRemove = {
                        recents.remove(it.uri)
                        files = recents.list()
                    },
                )
            }
        }
    }

    override fun onResume() {
        super.onResume()
        files = recents.list()
    }

    /** Fills the file just created at [uri] with a blank [format] file and opens it for editing. */
    private fun create(uri: Uri, format: DocFormat) {
        val blank = if (format == DocFormat.DOCX) R.raw.blank_document else R.raw.blank_spreadsheet
        thread {
            val ok = try {
                contentResolver.openOutputStream(uri, "w")!!.use { out ->
                    resources.openRawResource(blank).use { it.copyTo(out) }
                }
                true
            } catch (e: Exception) {
                false
            }
            runOnUiThread {
                if (ok) open(uri, format.mime, startEditing = true)
                else Toast.makeText(this, R.string.create_failed, Toast.LENGTH_LONG).show()
            }
        }
    }

    private fun open(uri: Uri, mime: String?, startEditing: Boolean = false) {
        val intent = Intent(this, ViewerActivity::class.java).setAction(Intent.ACTION_VIEW)
            .putExtra(ViewerActivity.EXTRA_START_EDITING, startEditing)
        if (RecentFiles.isOwnCopy(this, uri)) {
            // file:// URIs may not travel as intent data; the viewer rebuilds it from the path.
            intent.putExtra(ViewerActivity.EXTRA_OWN_COPY, uri.path).setType(mime)
        } else {
            intent.setDataAndType(uri, mime).addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        }
        try {
            startActivity(intent)
        } catch (e: SecurityException) {
            // Our access to the file has ended (it was lent by another app), so it can't be reopened.
            recents.remove(uri)
            files = recents.list()
            Toast.makeText(this, R.string.file_unavailable, Toast.LENGTH_LONG).show()
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
private fun HomeScreen(
    files: List<RecentFile>,
    onOpenPicker: () -> Unit,
    onNewDocument: () -> Unit,
    onNewSpreadsheet: () -> Unit,
    onOpen: (RecentFile) -> Unit,
    onRemove: (RecentFile) -> Unit,
) {
    Scaffold(
        topBar = { TopAppBar(title = { Text(stringResource(R.string.app_name)) }) },
        floatingActionButton = {
            Column(horizontalAlignment = Alignment.End) {
                NewFileButton(onNewDocument, onNewSpreadsheet)
                Spacer(Modifier.height(16.dp))
                ExtendedFloatingActionButton(onClick = onOpenPicker) { Text(stringResource(R.string.open_file)) }
            }
        },
    ) { padding ->
        if (files.isEmpty()) {
            Column(
                modifier = Modifier.fillMaxSize().padding(padding).padding(32.dp),
                verticalArrangement = Arrangement.Center,
                horizontalAlignment = Alignment.CenterHorizontally,
            ) {
                Text(stringResource(R.string.no_recent_title), style = MaterialTheme.typography.titleMedium)
                Spacer(Modifier.height(8.dp))
                Text(
                    stringResource(R.string.no_recent_body),
                    style = MaterialTheme.typography.bodyMedium,
                    color = MaterialTheme.colorScheme.onSurfaceVariant,
                    textAlign = TextAlign.Center,
                )
            }
        } else {
            LazyColumn(
                modifier = Modifier.fillMaxSize().padding(padding),
                // Room for the floating buttons below the last row.
                contentPadding = PaddingValues(bottom = 168.dp),
            ) {
                item {
                    Text(
                        stringResource(R.string.recent_files),
                        style = MaterialTheme.typography.labelLarge,
                        color = MaterialTheme.colorScheme.onSurfaceVariant,
                        modifier = Modifier.padding(horizontal = 16.dp, vertical = 8.dp),
                    )
                }
                items(files, key = { it.uri.toString() }) { file ->
                    RecentRow(file, onClick = { onOpen(file) }, onRemove = { onRemove(file) })
                }
            }
        }
    }
}

/** "+" with a menu: new document, new spreadsheet. */
@Composable
private fun NewFileButton(onNewDocument: () -> Unit, onNewSpreadsheet: () -> Unit) {
    var open by remember { mutableStateOf(false) }
    Box {
        FloatingActionButton(onClick = { open = true }) {
            Icon(Icons.Filled.Add, contentDescription = stringResource(R.string.new_file))
        }
        DropdownMenu(expanded = open, onDismissRequest = { open = false }) {
            DropdownMenuItem(
                text = { Text(stringResource(R.string.new_document)) },
                leadingIcon = { FormatBadge(DocFormat.DOCX, small = true) },
                onClick = {
                    open = false
                    onNewDocument()
                },
            )
            DropdownMenuItem(
                text = { Text(stringResource(R.string.new_spreadsheet)) },
                leadingIcon = { FormatBadge(DocFormat.XLSX, small = true) },
                onClick = {
                    open = false
                    onNewSpreadsheet()
                },
            )
        }
    }
}

/** The coloured file-type badge of the recent list. */
@Composable
private fun FormatBadge(format: DocFormat, small: Boolean = false) {
    Box(
        modifier = Modifier
            .size(width = if (small) 30.dp else 40.dp, height = if (small) 36.dp else 48.dp)
            .clip(RoundedCornerShape(if (small) 4.dp else 6.dp))
            .background(format.badgeColor()),
        contentAlignment = Alignment.Center,
    ) {
        Text(format.label, color = Color.White, fontSize = if (small) 10.sp else 12.sp, fontWeight = FontWeight.Bold)
    }
}

@Composable
private fun RecentRow(file: RecentFile, onClick: () -> Unit, onRemove: () -> Unit) {
    Row(
        modifier = Modifier
            .fillMaxWidth()
            .clickable(onClick = onClick)
            .padding(start = 16.dp, top = 8.dp, bottom = 8.dp, end = 4.dp),
        verticalAlignment = Alignment.CenterVertically,
    ) {
        FormatBadge(file.format)
        Spacer(Modifier.width(16.dp))
        Column(Modifier.weight(1f)) {
            Text(
                file.name,
                style = MaterialTheme.typography.bodyLarge,
                maxLines = 1,
                overflow = TextOverflow.Ellipsis,
            )
            Text(
                DateUtils.getRelativeTimeSpanString(
                    file.openedAt, System.currentTimeMillis(), DateUtils.MINUTE_IN_MILLIS,
                ).toString(),
                style = MaterialTheme.typography.bodySmall,
                color = MaterialTheme.colorScheme.onSurfaceVariant,
            )
        }
        IconButton(onClick = onRemove) {
            Icon(
                Icons.Filled.Close,
                contentDescription = stringResource(R.string.remove_from_recents),
                tint = MaterialTheme.colorScheme.onSurfaceVariant,
            )
        }
    }
}
