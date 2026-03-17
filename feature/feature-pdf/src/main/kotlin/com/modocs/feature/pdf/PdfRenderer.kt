package com.modocs.feature.pdf

import android.content.Context
import android.graphics.Bitmap
import android.graphics.pdf.PdfRenderer as AndroidPdfRenderer
import android.net.Uri
import android.os.ParcelFileDescriptor
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.sync.Mutex
import kotlinx.coroutines.sync.withLock
import kotlinx.coroutines.withContext

/**
 * Wrapper around Android's PdfRenderer that handles page rendering
 * with caching and thread safety.
 */
class PdfRendererWrapper private constructor(
    private val fileDescriptor: ParcelFileDescriptor,
    private val renderer: AndroidPdfRenderer,
) : AutoCloseable {

    val pageCount: Int get() = renderer.pageCount

    private val mutex = Mutex()

    // LRU-style cache: pageIndex -> rendered bitmap
    private val pageCache = LinkedHashMap<CacheKey, Bitmap>(16, 0.75f, true)
    private val maxCacheSize = 8

    data class CacheKey(val pageIndex: Int, val width: Int)

    /**
     * Render a page at a given width, maintaining aspect ratio.
     * Cached for repeated access (scrolling back and forth).
     */
    suspend fun renderPage(pageIndex: Int, renderWidth: Int): Bitmap? {
        if (pageIndex < 0 || pageIndex >= pageCount) return null

        val cacheKey = CacheKey(pageIndex, renderWidth)
        pageCache[cacheKey]?.let { return it }

        return withContext(Dispatchers.IO) {
            mutex.withLock {
                // Double-check after acquiring lock
                pageCache[cacheKey]?.let { return@withContext it }

                try {
                    val page = renderer.openPage(pageIndex)
                    val aspectRatio = page.height.toFloat() / page.width.toFloat()
                    val renderHeight = (renderWidth * aspectRatio).toInt()

                    val bitmap = Bitmap.createBitmap(
                        renderWidth,
                        renderHeight,
                        Bitmap.Config.ARGB_8888,
                    )
                    // White background
                    bitmap.eraseColor(android.graphics.Color.WHITE)

                    page.render(
                        bitmap,
                        null,
                        null,
                        AndroidPdfRenderer.Page.RENDER_MODE_FOR_DISPLAY,
                    )
                    page.close()

                    // Evict oldest if cache full
                    if (pageCache.size >= maxCacheSize) {
                        val oldestKey = pageCache.keys.first()
                        pageCache.remove(oldestKey)?.recycle()
                    }
                    pageCache[cacheKey] = bitmap

                    bitmap
                } catch (e: Exception) {
                    null
                }
            }
        }
    }

    /**
     * Get page dimensions without rendering (for layout calculations).
     */
    suspend fun getPageDimensions(pageIndex: Int): Pair<Int, Int>? {
        if (pageIndex < 0 || pageIndex >= pageCount) return null
        return withContext(Dispatchers.IO) {
            mutex.withLock {
                try {
                    val page = renderer.openPage(pageIndex)
                    val dims = page.width to page.height
                    page.close()
                    dims
                } catch (e: Exception) {
                    null
                }
            }
        }
    }

    override fun close() {
        pageCache.values.forEach { it.recycle() }
        pageCache.clear()
        try { renderer.close() } catch (_: Exception) {}
        try { fileDescriptor.close() } catch (_: Exception) {}
    }

    companion object {
        /**
         * Open a PDF from a content URI.
         */
        suspend fun open(context: Context, uri: Uri): PdfRendererWrapper? {
            return withContext(Dispatchers.IO) {
                try {
                    val fd = context.contentResolver.openFileDescriptor(uri, "r")
                        ?: return@withContext null
                    val renderer = AndroidPdfRenderer(fd)
                    PdfRendererWrapper(fd, renderer)
                } catch (e: Exception) {
                    null
                }
            }
        }
    }
}
