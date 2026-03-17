pluginManagement {
    repositories {
        google {
            content {
                includeGroupByRegex("com\\.android.*")
                includeGroupByRegex("com\\.google.*")
                includeGroupByRegex("androidx.*")
            }
        }
        mavenCentral()
        gradlePluginPortal()
    }
}

@Suppress("UnstableApiUsage")
dependencyResolutionManagement {
    repositoriesMode.set(RepositoriesMode.FAIL_ON_PROJECT_REPOS)
    repositories {
        google()
        mavenCentral()
    }
}

rootProject.name = "MoDocs"

include(":app")
include(":core:core-common")
include(":core:core-ui")
include(":core:core-model")
include(":core:core-storage")
include(":feature:feature-home")
include(":feature:feature-pdf")
include(":feature:feature-docx")
include(":feature:feature-xlsx")
include(":feature:feature-pptx")
include(":fonts")
