import java.util.Properties

plugins {
    alias(libs.plugins.android.application)
    alias(libs.plugins.kotlin.compose)
}

// Release signing comes from keystore.properties in the project root, which stays out of git
// (as do *.jks files). Without it, the release build is signed with the debug key and must not
// be published.
val keystoreFile = rootProject.file("keystore.properties")
val keystore = Properties().apply {
    if (keystoreFile.exists()) keystoreFile.inputStream().use { load(it) }
}

android {
    namespace = "com.officeeditor"
    compileSdk = 37

    defaultConfig {
        applicationId = "com.officeeditor"
        minSdk = 26
        targetSdk = 36
        versionCode = 3
        versionName = "1.1.0"
    }

    signingConfigs {
        if (keystoreFile.exists()) {
            create("release") {
                storeFile = rootProject.file(keystore.getProperty("storeFile"))
                storePassword = keystore.getProperty("storePassword")
                keyAlias = keystore.getProperty("keyAlias")
                keyPassword = keystore.getProperty("keyPassword")
            }
        }
    }

    buildTypes {
        release {
            isMinifyEnabled = true
            isShrinkResources = true
            proguardFiles(getDefaultProguardFile("proguard-android-optimize.txt"), "proguard-rules.pro")
            signingConfig = signingConfigs.findByName("release") ?: signingConfigs.getByName("debug")
        }
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }

    buildFeatures {
        compose = true
    }
}

// The viewers/editors are a Vite web app in /web, bundled into the APK's assets.
val webDir = rootProject.layout.projectDirectory.dir("web")
val buildWeb by tasks.registering(Exec::class) {
    description = "Builds the web viewers into app/src/main/assets/web"
    workingDir = webDir.asFile
    commandLine("npm", "run", "build")
    inputs.dir(webDir.dir("src"))
    inputs.dir(webDir.dir("scripts"))
    inputs.files(
        webDir.file("index.html"),
        webDir.file("package.json"),
        webDir.file("package-lock.json"),
        webDir.file("vite.config.ts"),
        webDir.file("tsconfig.json"),
    )
    outputs.dir(layout.projectDirectory.dir("src/main/assets/web"))
}
tasks.named("preBuild") { dependsOn(buildWeb) }

dependencies {
    implementation(platform(libs.compose.bom))
    implementation(libs.compose.ui)
    implementation(libs.compose.ui.tooling.preview)
    implementation(libs.compose.material3)
    implementation(libs.compose.material.icons)
    debugImplementation(libs.compose.ui.tooling)
    implementation(libs.activity.compose)
    implementation(libs.webkit)

    testImplementation(libs.junit)
}
