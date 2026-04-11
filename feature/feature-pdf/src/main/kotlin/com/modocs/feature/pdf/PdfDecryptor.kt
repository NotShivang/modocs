package com.modocs.feature.pdf

import android.content.Context
import android.net.Uri
import android.os.ParcelFileDescriptor
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.pdmodel.PDDocument
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.File
import java.io.IOException

/**
 * Decrypts password-protected PDF files using PDFBox.
 *
 * Loads the encrypted PDF with the user's password, removes all security,
 * and saves a decrypted copy to a temp file. The temp file can then be
 * opened by Android's built-in PdfRenderer.
 */
object PdfDecryptor {

    sealed class DecryptResult {
        data class Success(val tempFile: File) : DecryptResult()
        data object WrongPassword : DecryptResult()
        data class Failed(val message: String) : DecryptResult()
    }

    suspend fun decrypt(context: Context, uri: Uri, password: String): DecryptResult =
        withContext(Dispatchers.IO) {
            PDFBoxResourceLoader.init(context)

            val inputStream = context.contentResolver.openInputStream(uri)
                ?: return@withContext DecryptResult.Failed("Cannot open file")

            try {
                val document = PDDocument.load(inputStream, password)

                val tempFile = File(context.cacheDir, "decrypted_${System.nanoTime()}.pdf")
                document.isAllSecurityToBeRemoved = true
                document.save(tempFile)
                document.close()

                DecryptResult.Success(tempFile)
            } catch (_: IOException) {
                // PDFBox throws IOException with message containing "password" for wrong passwords
                DecryptResult.WrongPassword
            } catch (e: Exception) {
                DecryptResult.Failed(e.message ?: "Decryption failed")
            } finally {
                inputStream.close()
            }
        }
}
