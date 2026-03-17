package com.modocs.core.storage.db

import androidx.room.Entity
import androidx.room.PrimaryKey

@Entity(tableName = "recent_files")
data class RecentFileEntity(
    @PrimaryKey(autoGenerate = true) val id: Long = 0,
    val uri: String,
    val displayName: String,
    val documentType: String,
    val lastOpenedTimestamp: Long,
    val fileSizeBytes: Long?,
)
