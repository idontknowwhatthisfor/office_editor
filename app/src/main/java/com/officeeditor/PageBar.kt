package com.officeeditor

import androidx.activity.compose.BackHandler
import androidx.compose.foundation.border
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.widthIn
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.BasicTextField
import androidx.compose.foundation.text.KeyboardActions
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Close
import androidx.compose.material.icons.filled.Edit
import androidx.compose.material.icons.filled.KeyboardArrowDown
import androidx.compose.material.icons.filled.KeyboardArrowUp
import androidx.compose.material.icons.filled.Search
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.saveable.rememberSaveable
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.focus.FocusRequester
import androidx.compose.ui.focus.focusRequester
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.SolidColor
import androidx.compose.ui.graphics.vector.ImageVector
import androidx.compose.ui.graphics.vector.PathParser
import androidx.compose.ui.platform.LocalConfiguration
import androidx.compose.ui.platform.LocalSoftwareKeyboardController
import androidx.compose.ui.res.stringResource
import androidx.compose.ui.text.input.ImeAction
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import kotlinx.coroutines.delay
import org.json.JSONObject

/**
 * What the PDF and Word viewers' bottom bar shows; sent by the page (`setPageBar`, pagebar.ts).
 * [page] is 0-based. [shown] is false while the page shows a toolbar of its own (edit mode).
 */
data class PageBarState(
    val page: Int,
    val count: Int,
    val query: String,
    val current: Int,
    val total: Int,
    val searching: Boolean,
    val shown: Boolean,
) {
    companion object {
        fun parse(json: String): PageBarState? = try {
            val o = JSONObject(json)
            PageBarState(
                page = o.getInt("page"),
                count = o.getInt("count"),
                query = o.optString("query"),
                current = o.optInt("current", -1),
                total = o.optInt("total"),
                searching = o.optBoolean("searching"),
                shown = o.optBoolean("shown", true),
            )
        } catch (e: Exception) {
            null
        }
    }
}

/**
 * The bottom bar of the PDF and Word viewers: thumbnails, page indicator, search, edit. Drawn
 * here rather than by the page so it stays put while the page is pinch-zoomed and panned (a page
 * element can only follow the zoomed view a frame late). Every button is passed on to the page
 * through [send] (`window.officePageBar(action, arg)`); the thumbnail strip is the page's own and
 * opens above this bar. Same height as the page's spacer row (pagebar.css): 56 dp, 44 on a short
 * screen.
 */
@Composable
fun PageBar(state: PageBarState, send: (action: String, arg: Any?) -> Unit, modifier: Modifier = Modifier) {
    var searching by rememberSaveable { mutableStateOf(false) }
    var text by rememberSaveable { mutableStateOf("") }
    // The text the page last searched for, so a typing pause and Enter don't both search.
    val sent = remember { arrayOf<String?>(null) }
    val keyboard = LocalSoftwareKeyboardController.current
    val focus = remember { FocusRequester() }
    val short = LocalConfiguration.current.screenHeightDp < 480

    fun find() {
        if (sent[0] == text) return
        sent[0] = text
        send("find", text)
    }
    fun close() {
        searching = false
        text = ""
        sent[0] = null
        send("clear", null)
    }

    LaunchedEffect(text, searching) {
        if (!searching) return@LaunchedEffect
        delay(250)
        find()
    }
    LaunchedEffect(searching) {
        if (searching) focus.requestFocus()
    }
    BackHandler(enabled = searching) { close() }

    Surface(modifier.fillMaxWidth(), color = MaterialTheme.colorScheme.surface) {
        Column {
            HorizontalDivider(color = MaterialTheme.colorScheme.outlineVariant)
            Box(Modifier.fillMaxWidth().height(if (short) 44.dp else 56.dp), contentAlignment = Alignment.Center) {
                // On a wide screen the buttons stay together in the middle (as the page's bars do).
                Row(
                    Modifier.widthIn(max = 600.dp).fillMaxWidth().padding(horizontal = 4.dp),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically,
                ) {
                    if (!searching) {
                        IconButton(onClick = { send("pages", null) }) {
                            Icon(PagesIcon, stringResource(R.string.pages))
                        }
                        TextButton(onClick = { send("pages", null) }) {
                            Text(
                                "${state.page + 1} / ${state.count}",
                                color = MaterialTheme.colorScheme.onSurface,
                                style = MaterialTheme.typography.bodyLarge.copy(fontFeatureSettings = "tnum"),
                            )
                        }
                        IconButton(onClick = { searching = true }) {
                            Icon(Icons.Filled.Search, stringResource(R.string.search))
                        }
                        IconButton(onClick = { send("edit", null) }) {
                            Icon(Icons.Filled.Edit, stringResource(R.string.edit))
                        }
                    } else {
                        SearchField(
                            text,
                            onChange = { text = it },
                            onSearch = {
                                keyboard?.hide() // so the result is visible
                                find()
                                send("step", 1)
                            },
                            Modifier.weight(1f).padding(start = 8.dp).focusRequester(focus),
                        )
                        Text(
                            when {
                                state.total > 0 -> "${state.current + 1}/${state.total}"
                                state.query.isEmpty() -> ""
                                state.searching -> "…"
                                else -> stringResource(R.string.no_results)
                            },
                            Modifier.padding(horizontal = 4.dp),
                            color = MaterialTheme.colorScheme.onSurfaceVariant,
                            fontSize = 14.sp,
                            maxLines = 1,
                        )
                        IconButton(onClick = { send("step", -1) }) {
                            Icon(Icons.Filled.KeyboardArrowUp, stringResource(R.string.previous_result))
                        }
                        IconButton(onClick = { send("step", 1) }) {
                            Icon(Icons.Filled.KeyboardArrowDown, stringResource(R.string.next_result))
                        }
                        IconButton(onClick = { close() }) {
                            Icon(Icons.Filled.Close, stringResource(R.string.close))
                        }
                    }
                }
            }
        }
    }
}

@Composable
private fun SearchField(text: String, onChange: (String) -> Unit, onSearch: () -> Unit, modifier: Modifier) {
    val colors = MaterialTheme.colorScheme
    BasicTextField(
        value = text,
        onValueChange = onChange,
        modifier = modifier
            .height(40.dp)
            .border(1.dp, colors.outline, RoundedCornerShape(20.dp)),
        singleLine = true,
        textStyle = MaterialTheme.typography.bodyLarge.copy(color = colors.onSurface),
        cursorBrush = SolidColor(colors.primary),
        keyboardOptions = KeyboardOptions(imeAction = ImeAction.Search),
        keyboardActions = KeyboardActions(onSearch = { onSearch() }),
        decorationBox = { field ->
            Box(Modifier.padding(horizontal = 12.dp), contentAlignment = Alignment.CenterStart) {
                if (text.isEmpty()) {
                    Text(
                        stringResource(R.string.search_hint),
                        color = colors.onSurfaceVariant,
                        style = MaterialTheme.typography.bodyLarge,
                        maxLines = 1,
                        overflow = TextOverflow.Ellipsis,
                    )
                }
                field()
            }
        },
    )
}

/** Material Symbols "apps" (Apache-2.0), as the page's bar had it. */
private val PagesIcon: ImageVector = ImageVector.Builder("pages", 24.dp, 24.dp, 24f, 24f).addPath(
    PathParser().parsePathString(
        "M4 8h4V4H4v4zm6 12h4v-4h-4v4zm-6 0h4v-4H4v4zm0-6h4v-4H4v4zm6 0h4v-4h-4v4zm6-10v4h4V4h-4zm-6 4h4V4h-4v4zm6 6h4v-4h-4v4zm0 6h4v-4h-4v4z",
    ).toNodes(),
    fill = SolidColor(Color.Black),
).build()
