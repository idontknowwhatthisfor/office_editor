package com.officeeditor

import androidx.compose.foundation.ExperimentalFoundationApi
import androidx.compose.foundation.Canvas
import androidx.compose.foundation.background
import androidx.compose.foundation.combinedClickable
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.widthIn
import androidx.compose.foundation.shape.CircleShape
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.material3.ripple
import androidx.compose.runtime.Composable
import androidx.compose.runtime.remember
import androidx.compose.foundation.interaction.MutableInteractionSource
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.alpha
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Matrix
import androidx.compose.ui.graphics.Path
import androidx.compose.ui.graphics.drawscope.scale
import androidx.compose.ui.graphics.vector.PathParser
import androidx.compose.ui.semantics.contentDescription
import androidx.compose.ui.semantics.semantics
import androidx.compose.ui.text.font.FontStyle
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.Dp
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import org.json.JSONArray
import org.json.JSONObject

/** One button of an editor toolbar, as the page describes it (toolbar-mirror.ts). */
class ToolButton(
    val label: String,
    val active: Boolean,
    val disabled: Boolean,
    /** Icon size in dp: the box its 24-unit viewBox is drawn in. */
    val size: Float,
    /** The icon's paths, already moved by their groups' transforms; empty for a text button. */
    val paths: List<Path>,
    val text: String,
    val bold: Boolean,
    val italic: Boolean,
)

/** An editor toolbar the shell draws (`setToolbar`); [height] in dp, as the page's row. */
class ToolbarState(val height: Float, val buttons: List<ToolButton>) {
    companion object {
        /** Null for '' (no toolbar) or anything unreadable. */
        fun parse(json: String): ToolbarState? = try {
            if (json.isEmpty()) null else {
                val o = JSONObject(json)
                val list = o.getJSONArray("buttons")
                ToolbarState(o.getDouble("height").toFloat(), (0 until list.length()).map { button(list.getJSONObject(it)) })
            }
        } catch (e: Exception) {
            null
        }

        private fun button(b: JSONObject) = ToolButton(
            label = b.optString("label"),
            active = b.optBoolean("active"),
            disabled = b.optBoolean("disabled"),
            size = b.optDouble("size", 24.0).toFloat(),
            paths = b.optJSONArray("paths")?.let { paths -> (0 until paths.length()).map { path(paths.getJSONObject(it)) } }.orEmpty(),
            text = b.optString("text"),
            bold = b.optBoolean("bold"),
            italic = b.optBoolean("italic"),
        )

        private fun path(p: JSONObject): Path {
            val path = PathParser().parsePathString(p.getString("d")).toPath()
            val m: JSONArray? = p.optJSONArray("m")
            if (m != null && m.length() == 6) {
                // An SVG affine [a b c d e f], as a 4×4 matrix.
                val matrix = Matrix()
                matrix.values[Matrix.ScaleX] = m.getDouble(0).toFloat()
                matrix.values[Matrix.SkewY] = m.getDouble(1).toFloat()
                matrix.values[Matrix.SkewX] = m.getDouble(2).toFloat()
                matrix.values[Matrix.ScaleY] = m.getDouble(3).toFloat()
                matrix.values[Matrix.TranslateX] = m.getDouble(4).toFloat()
                matrix.values[Matrix.TranslateY] = m.getDouble(5).toFloat()
                path.transform(matrix)
            }
            return path
        }
    }
}

/**
 * The PDF and Word editors' toolbar, drawn natively over the bottom of the WebView so it stays put
 * while the page is pinch-zoomed (a page element can only follow the zoomed view a frame late).
 * The page keeps its own toolbar, unseen, and describes it ([ToolbarState]); taps and holds go back
 * through [send] (`window.officeToolbar(index, kind)`) and are replayed on the page's buttons, so the
 * editors' behaviour stays all in the page. Looks like the page's toolbar: buttons share the width
 * (within 600 dp on a wide screen), an active one is tinted, a disabled one faded.
 */
@OptIn(ExperimentalFoundationApi::class)
@Composable
fun MirroredToolbar(state: ToolbarState, send: (index: Int, kind: String) -> Unit, modifier: Modifier = Modifier) {
    val colors = MaterialTheme.colorScheme
    val buttonHeight: Dp = minOf(48f, state.height - 8f).coerceAtLeast(32f).dp
    Surface(modifier.fillMaxWidth(), color = colors.surface) {
        Column {
            HorizontalDivider(color = colors.outlineVariant)
            Box(Modifier.fillMaxWidth().height((state.height - 1).coerceAtLeast(32f).dp), contentAlignment = Alignment.Center) {
                Row(
                    Modifier.widthIn(max = 600.dp).fillMaxWidth().padding(horizontal = 4.dp),
                    verticalAlignment = Alignment.CenterVertically,
                ) {
                    state.buttons.forEachIndexed { index, b ->
                        val tint = if (b.active) colors.primary else colors.onSurface
                        Box(
                            Modifier
                                .weight(1f)
                                .height(buttonHeight)
                                .clip(CircleShape)
                                .background(if (b.active) colors.primary.copy(alpha = 0.22f) else colors.surface)
                                .combinedClickable(
                                    interactionSource = remember { MutableInteractionSource() },
                                    indication = ripple(),
                                    enabled = !b.disabled,
                                    onClick = { send(index, "tap") },
                                    onLongClick = { send(index, "hold") },
                                )
                                .semantics { contentDescription = b.label }
                                .alpha(if (b.disabled) 0.35f else 1f),
                            contentAlignment = Alignment.Center,
                        ) {
                            if (b.paths.isNotEmpty()) {
                                Canvas(Modifier.size(b.size.dp)) {
                                    scale(size.width / 24f, pivot = androidx.compose.ui.geometry.Offset.Zero) {
                                        for (path in b.paths) drawPath(path, tint)
                                    }
                                }
                            } else {
                                Text(
                                    b.text,
                                    color = tint,
                                    fontSize = 18.sp,
                                    fontWeight = if (b.bold) FontWeight.Bold else FontWeight.Normal,
                                    fontStyle = if (b.italic) FontStyle.Italic else FontStyle.Normal,
                                )
                            }
                        }
                    }
                }
            }
        }
    }
}
