package com.modocs.core.storage.di

import android.content.Context
import androidx.room.Room
import com.modocs.core.storage.db.MoDocsDatabase
import com.modocs.core.storage.db.RecentFileDao
import dagger.Module
import dagger.Provides
import dagger.hilt.InstallIn
import dagger.hilt.android.qualifiers.ApplicationContext
import dagger.hilt.components.SingletonComponent
import javax.inject.Singleton

@Module
@InstallIn(SingletonComponent::class)
object StorageModule {

    @Provides
    @Singleton
    fun provideDatabase(@ApplicationContext context: Context): MoDocsDatabase {
        return Room.databaseBuilder(
            context,
            MoDocsDatabase::class.java,
            "modocs.db",
        ).build()
    }

    @Provides
    fun provideRecentFileDao(database: MoDocsDatabase): RecentFileDao {
        return database.recentFileDao()
    }
}
