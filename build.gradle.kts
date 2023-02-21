import com.github.jengelman.gradle.plugins.shadow.tasks.ShadowJar

val attendanceMain = "alpha.sig.AttendanceKt"

plugins {
    val kotlinVersion = "1.7.20"
    kotlin("jvm") version kotlinVersion
    id("application")
    id("com.github.johnrengelman.shadow") version "5.2.0"
}

group = "alpha.sig"
version = "1.0-SNAPSHOT"

description = """Attendance Bot"""


repositories {
    mavenCentral()
}

buildscript {
    repositories {
        mavenCentral()
    }
}

dependencies {
    @Suppress("VulnerableLibrariesLocal", "RedundantSuppression")
    implementation("org.docx4j:docx4j-JAXB-ReferenceImpl:11.4.8") {
        // This version has a vulnerability because they use such an old version, so I updated it below.
        exclude("commons-codec:commons-codec:*")
    }
    implementation("commons-codec:commons-codec:1.15")
    implementation(kotlin("stdlib-jdk8"))
}

sourceSets.main {
    kotlin {
        srcDirs("src/kotlin")
    }
    resources.srcDir("resources")
}

sourceSets.test {
    kotlin {
        srcDirs("test/kotlin")
    }
}

tasks.compileKotlin {
    kotlinOptions {
        jvmTarget = "1.8"
    }
}

application {
    mainClass.set(attendanceMain)
    // Shadow requires us to set the main class name this way, which is dumb.
    @Suppress("DEPRECATION")
    mainClassName = attendanceMain
}

tasks.withType<ShadowJar> {
    manifest {
        attributes("Main-Class" to application.mainClass.get())
    }
    minimize { // This saves us a whopping 400 KiB :O
        exclude(dependency("org.docx4j:docx4j-JAXB-ReferenceImpl:.*"))
    }
}

tasks.named<JavaExec>("run") {
    standardInput = System.`in`
}

tasks.wrapper {
    gradleVersion = "7.0"
}
