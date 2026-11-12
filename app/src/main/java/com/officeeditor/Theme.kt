package com.officeeditor

import android.os.Build
import androidx.compose.foundation.isSystemInDarkTheme
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.darkColorScheme
import androidx.compose.material3.dynamicDarkColorScheme
import androidx.compose.material3.dynamicLightColorScheme
import androidx.compose.material3.lightColorScheme
import androidx.compose.runtime.Composable
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext

private val Blue = Color(0xFF2F5BD7)

@Composable
fun OfficeEditorTheme(content: @Composable () -> Unit) {
    val dark = isSystemInDarkTheme()
    val context = LocalContext.current
    val colors = when {
        Build.VERSION.SDK_INT >= Build.VERSION_CODES.S ->
            if (dark) dynamicDarkColorScheme(context) else dynamicLightColorScheme(context)
        dark -> darkColorScheme(primary = Color(0xFFB5C4FF))
        else -> lightColorScheme(primary = Blue)
    }
    MaterialTheme(colorScheme = colors, content = content)
}

/** The badge colour each format is conventionally known by. */
fun DocFormat.badgeColor(): Color = when (this) {
    DocFormat.PDF -> Color(0xFFD93025)
    DocFormat.DOCX -> Color(0xFF2B579A)
    DocFormat.XLSX -> Color(0xFF1E7145)
    DocFormat.PPTX -> Color(0xFFC43E1C)
}
