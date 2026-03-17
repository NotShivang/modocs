package com.modocs.fonts

import android.content.Context
import androidx.compose.ui.text.font.Font
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontStyle
import androidx.compose.ui.text.font.FontWeight

/**
 * Resolves font names from documents (e.g., "Calibri", "Arial") to bundled
 * metric-compatible open-source font families.
 *
 * To set up fonts, download the following and place them in fonts/src/main/assets/fonts/:
 * - Carlito (replaces Calibri): https://fonts.google.com/specimen/Carlito
 * - Caladea (replaces Cambria): https://fonts.google.com/specimen/Caladea
 * - Liberation Sans (replaces Arial): https://github.com/liberationfonts/liberation-fonts
 * - Liberation Serif (replaces Times New Roman): same repo
 * - Liberation Mono (replaces Courier New): same repo
 */
class FontResolver(private val context: Context) {

    private val fontCache = mutableMapOf<String, FontFamily>()

    // Maps Office font names (lowercase) to asset paths.
    // Each entry maps to [Regular, Bold, Italic, BoldItalic] file names.
    private val fontMapping = mapOf(
        "calibri" to FontFiles("Carlito"),
        "calibri light" to FontFiles("Carlito"),
        "cambria" to FontFiles("Caladea"),
        "arial" to FontFiles("LiberationSans"),
        "helvetica" to FontFiles("LiberationSans"),
        "times new roman" to FontFiles("LiberationSerif"),
        "times" to FontFiles("LiberationSerif"),
        "courier new" to FontFiles("LiberationMono"),
        "courier" to FontFiles("LiberationMono"),
    )

    fun resolve(fontName: String?): FontFamily {
        val key = fontName?.lowercase()?.trim() ?: "calibri"
        return fontCache.getOrPut(key) {
            val files = fontMapping[key] ?: fontMapping["calibri"]!!
            createFontFamily(files)
        }
    }

    private fun createFontFamily(files: FontFiles): FontFamily {
        val fonts = mutableListOf<Font>()

        tryLoadFont(files.regular, FontWeight.Normal, FontStyle.Normal)?.let { fonts.add(it) }
        tryLoadFont(files.bold, FontWeight.Bold, FontStyle.Normal)?.let { fonts.add(it) }
        tryLoadFont(files.italic, FontWeight.Normal, FontStyle.Italic)?.let { fonts.add(it) }
        tryLoadFont(files.boldItalic, FontWeight.Bold, FontStyle.Italic)?.let { fonts.add(it) }

        return if (fonts.isNotEmpty()) FontFamily(fonts) else FontFamily.Default
    }

    private fun tryLoadFont(assetPath: String, weight: FontWeight, style: FontStyle): Font? {
        return try {
            // Verify the asset exists
            context.assets.open(assetPath).close()
            Font(assetPath, context.assets, weight, style)
        } catch (_: Exception) {
            null
        }
    }

    private data class FontFiles(val baseName: String) {
        val regular = "fonts/$baseName-Regular.ttf"
        val bold = "fonts/$baseName-Bold.ttf"
        val italic = "fonts/$baseName-Italic.ttf"
        val boldItalic = "fonts/$baseName-BoldItalic.ttf"
    }
}
