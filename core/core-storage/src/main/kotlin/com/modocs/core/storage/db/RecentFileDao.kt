package com.modocs.core.storage.db

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import kotlinx.coroutines.flow.Flow

@Dao
interface RecentFileDao {

    @Query("SELECT * FROM recent_files ORDER BY lastOpenedTimestamp DESC LIMIT 50")
    fun getRecentFiles(): Flow<List<RecentFileEntity>>

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun upsert(entity: RecentFileEntity)

    @Query("SELECT * FROM recent_files WHERE uri = :uri LIMIT 1")
    suspend fun findByUri(uri: String): RecentFileEntity?

    @Query("DELETE FROM recent_files WHERE id = :id")
    suspend fun deleteById(id: Long)

    @Query("DELETE FROM recent_files")
    suspend fun deleteAll()
}
