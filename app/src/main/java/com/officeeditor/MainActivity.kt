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
import androidx.compose.material3.ExtendedFloatingActionButton
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.material3.TopAppBar
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
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

/** Home: recent files and "Open file". */
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
                HomeScreen(
                    files = files,
                    onOpenPicker = { picker.launch(DocFormat.mimeTypes) },
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

    private fun open(uri: Uri, mime: String?) {
        val intent = Intent(this, ViewerActivity::class.java).setAction(Intent.ACTION_VIEW)
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
    onOpen: (RecentFile) -> Unit,
    onRemove: (RecentFile) -> Unit,
) {
    Scaffold(
        topBar = { TopAppBar(title = { Text(stringResource(R.string.app_name)) }) },
        floatingActionButton = {
            ExtendedFloatingActionButton(
                onClick = onOpenPicker,
                icon = { Icon(Icons.Filled.Add, contentDescription = null) },
                text = { Text(stringResource(R.string.open_file)) },
            )
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
                // Room for the floating button below the last row.
                contentPadding = PaddingValues(bottom = 96.dp),
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

@Composable
private fun RecentRow(file: RecentFile, onClick: () -> Unit, onRemove: () -> Unit) {
    Row(
        modifier = Modifier
            .fillMaxWidth()
            .clickable(onClick = onClick)
            .padding(start = 16.dp, top = 8.dp, bottom = 8.dp, end = 4.dp),
        verticalAlignment = Alignment.CenterVertically,
    ) {
        Box(
            modifier = Modifier
                .size(width = 40.dp, height = 48.dp)
                .clip(RoundedCornerShape(6.dp))
                .background(file.format.badgeColor()),
            contentAlignment = Alignment.Center,
        ) {
            Text(file.format.label, color = Color.White, fontSize = 12.sp, fontWeight = FontWeight.Bold)
        }
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
