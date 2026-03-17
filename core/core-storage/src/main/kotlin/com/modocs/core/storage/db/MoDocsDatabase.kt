package com.modocs.core.storage.db

import androidx.room.Database
import androidx.room.RoomDatabase

@Database(
    entities = [RecentFileEntity::class],
    version = 1,
    exportSchema = false,
)
abstract class MoDocsDatabase : RoomDatabase() {
    abstract fun recentFileDao(): RecentFileDao
}
