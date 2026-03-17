package com.modocs.feature.pdf

import androidx.compose.ui.geometry.Offset

/**
 * Annotations that users can place on PDF pages for fill & sign.
 * Coordinates are normalized (0..1) relative to page dimensions.
 */
sealed interface PdfAnnotation {
    val id: String
    val pageIndex: Int
    /** Normalized x position (0..1). */
    val x: Float
    /** Normalized y position (0..1). */
    val y: Float
}

data class TextAnnotation(
    override val id: String,
    override val pageIndex: Int,
    override val x: Float,
    override val y: Float,
    val text: String,
    val fontSizeSp: Float = 18f,
) : PdfAnnotation

data class SignatureAnnotation(
    override val id: String,
    override val pageIndex: Int,
    override val x: Float,
    override val y: Float,
    /** Normalized width (0..1). */
    val width: Float = 0.3f,
    /** Normalized height (0..1). */
    val height: Float = 0.08f,
    /** Signature stroke paths — list of strokes, each stroke is a list of normalized points. */
    val strokes: List<List<Offset>>,
) : PdfAnnotation

data class CheckmarkAnnotation(
    override val id: String,
    override val pageIndex: Int,
    override val x: Float,
    override val y: Float,
    val sizeSp: Float = 20f,
    /** true = checkmark, false = X mark. */
    val isCheck: Boolean = true,
) : PdfAnnotation

data class DateAnnotation(
    override val id: String,
    override val pageIndex: Int,
    override val x: Float,
    override val y: Float,
    val dateText: String,
    val fontSizeSp: Float = 18f,
) : PdfAnnotation

enum class FillSignTool {
    TEXT,
    SIGNATURE,
    CHECKMARK,
    CROSS,
    DATE,
}
