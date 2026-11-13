package com.officeeditor

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.BackHandler
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.PaddingValues
import androidx.compose.foundation.layout.WindowInsets
import androidx.compose.foundation.layout.WindowInsetsSides
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.only
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.safeDrawing
import androidx.compose.foundation.layout.widthIn
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.HorizontalDivider
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
import androidx.compose.ui.res.stringResource
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import org.json.JSONArray

/** A bundled third-party component and its licence (from assets/web/licences.json). */
private data class Component(val name: String, val version: String, val licence: String, val usedFor: String, val text: String)

/**
 * About & licences: the app's version and privacy promise, and the licence of every bundled
 * third-party component, as those licences require. The list is built from the installed
 * packages at build time by web/scripts/collect-licences.mjs.
 */
class AboutActivity : ComponentActivity() {
    @OptIn(ExperimentalMaterial3Api::class)
    override fun onCreate(savedInstanceState: Bundle?) {
        enableEdgeToEdge()
        super.onCreate(savedInstanceState)
        val components = readComponents()
        val version = packageManager.getPackageInfo(packageName, 0).versionName.orEmpty()

        setContent {
            OfficeEditorTheme {
                var open by remember { mutableStateOf<Component?>(null) }
                BackHandler(enabled = open != null) { open = null }
                Scaffold(
                    topBar = {
                        TopAppBar(
                            title = {
                                Text(
                                    open?.name ?: stringResource(R.string.about),
                                    maxLines = 1,
                                    overflow = TextOverflow.Ellipsis,
                                )
                            },
                            expandedHeight = appBarHeight(),
                            windowInsets = WindowInsets.safeDrawing.only(WindowInsetsSides.Horizontal + WindowInsetsSides.Top),
                            navigationIcon = {
                                IconButton(onClick = { if (open != null) open = null else finish() }) {
                                    Icon(Icons.AutoMirrored.Filled.ArrowBack, stringResource(R.string.back))
                                }
                            },
                        )
                    },
                    contentWindowInsets = WindowInsets.safeDrawing,
                ) { padding ->
                    Box(Modifier.fillMaxSize().padding(padding), contentAlignment = Alignment.TopCenter) {
                        val shown = open
                        if (shown != null) LicenceText(shown) else Overview(version, components) { open = it }
                    }
                }
            }
        }
    }

    private fun readComponents(): List<Component> = try {
        val array = JSONArray(assets.open("web/licences.json").bufferedReader().use { it.readText() })
        (0 until array.length()).map { i ->
            val o = array.getJSONObject(i)
            Component(o.getString("name"), o.getString("version"), o.getString("licence"), o.getString("usedFor"), o.getString("text"))
        }
    } catch (e: Exception) {
        emptyList()
    }
}

@Composable
private fun Overview(version: String, components: List<Component>, onOpen: (Component) -> Unit) {
    LazyColumn(
        modifier = Modifier.widthIn(max = 720.dp).fillMaxWidth(),
        contentPadding = PaddingValues(bottom = 24.dp),
    ) {
        item {
            Column(Modifier.padding(16.dp)) {
                Text(stringResource(R.string.app_name), style = MaterialTheme.typography.headlineSmall)
                Text(
                    stringResource(R.string.about_version, version),
                    style = MaterialTheme.typography.bodyMedium,
                    color = MaterialTheme.colorScheme.onSurfaceVariant,
                )
                Text(
                    stringResource(R.string.about_privacy),
                    style = MaterialTheme.typography.bodyMedium,
                    modifier = Modifier.padding(top = 12.dp),
                )
            }
            HorizontalDivider()
            Text(
                stringResource(R.string.open_source_licences),
                style = MaterialTheme.typography.titleMedium,
                modifier = Modifier.padding(start = 16.dp, end = 16.dp, top = 16.dp, bottom = 4.dp),
            )
            Text(
                stringResource(R.string.open_source_licences_body),
                style = MaterialTheme.typography.bodySmall,
                color = MaterialTheme.colorScheme.onSurfaceVariant,
                modifier = Modifier.padding(horizontal = 16.dp, vertical = 4.dp),
            )
        }
        items(components) { c ->
            Column(
                Modifier
                    .fillMaxWidth()
                    .clickable { onOpen(c) }
                    .padding(horizontal = 16.dp, vertical = 10.dp),
            ) {
                Text(if (c.version.isEmpty()) c.name else "${c.name} ${c.version}", style = MaterialTheme.typography.bodyLarge)
                Text(
                    "${c.licence} · ${c.usedFor}",
                    style = MaterialTheme.typography.bodySmall,
                    color = MaterialTheme.colorScheme.onSurfaceVariant,
                )
            }
        }
    }
}

/** The full licence text, in a fixed-width font so indented clauses stay aligned. */
@Composable
private fun LicenceText(c: Component) {
    Column(
        Modifier
            .widthIn(max = 720.dp)
            .fillMaxSize()
            .verticalScroll(rememberScrollState())
            .padding(16.dp),
    ) {
        Text(
            "${c.licence} · ${c.usedFor}",
            style = MaterialTheme.typography.bodyMedium,
            color = MaterialTheme.colorScheme.onSurfaceVariant,
            modifier = Modifier.padding(bottom = 12.dp),
        )
        Text(c.text, fontFamily = FontFamily.Monospace, fontSize = 11.sp, lineHeight = 15.sp)
    }
}
