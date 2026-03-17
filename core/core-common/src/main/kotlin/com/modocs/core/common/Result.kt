package com.modocs.core.common

sealed interface MoDocsResult<out T> {
    data class Success<T>(val data: T) : MoDocsResult<T>
    data class Error(val message: String, val cause: Throwable? = null) : MoDocsResult<Nothing>
    data object Loading : MoDocsResult<Nothing>
}
