package com.modocs.core.model

import com.modocs.core.common.DocumentType

data class RecentFile(
    val id: Long = 0,
    val uri: String,
    val displayName: String,
    val documentType: DocumentType,
    val lastOpenedTimestamp: Long,
    val fileSizeBytes: Long? = null,
)
