package alpha.sig

import com.squareup.moshi.JsonDataException
import com.squareup.moshi.Moshi
import jakarta.xml.bind.JAXBElement
import org.xlsx4j.sml.CTRst
import org.xlsx4j.sml.CTXstringWhitespace
import java.time.LocalDate

// This file is where all the cute little utility functions and extensions go.

@Suppress("UNCHECKED_CAST")
fun <T> Any.unwrap() = ((this as JAXBElement<*>).value) as T

fun <T> List<T>.subList(startIndex: Int) = subList(startIndex, size) // Why doesn't this just exist, Java?

/**
 * Creates all the wrapper objects necessary to create a properly formatted inline string.
 */
fun inlineStr(s: String): CTRst {
    val ctrst = CTRst()
    val strWhitespace = CTXstringWhitespace()
    strWhitespace.value = s
    ctrst.t = strWhitespace
    return ctrst
}

/**
 * Convert Microsoft's weird date format from Lotus 1-2-3 (Offset from Jan 1st 1900 w/leap days on centuries) to a [LocalDate].
 */
fun microsoftDateToLocalDate(microsoftDate: Long): LocalDate {
    val dateIfLotusDidntHaveBug = microsoftEpoch.plusDays(microsoftDate)
    return dateIfLotusDidntHaveBug.minusDays(((dateIfLotusDidntHaveBug.year / 100) - 18).toLong())
}

/**
 * Deserializes a JSON string into the type specified in [T].
 */
inline fun <reified T: Any> Moshi.fromJson(jsonStr: String): T =
    this.adapter(T::class.java)?.fromJson(jsonStr)
        ?: throw JsonDataException("Is there an adapter missing for ${T::class.simpleName}?")

fun <T> Moshi.fromJson(jsonStr: String, cls: Class<T>): T = this.adapter(cls)?.fromJson(jsonStr)
    ?: throw JsonDataException("Is there an adapter missing for ${cls.simpleName}?")

/**
 * Serializes a variable of type [T] into a JSON string. Will serialize null values if they are present.
 */
inline fun <reified T: Any> Moshi.toJson(obj: T, prettyPrint: Boolean = false): String =
    this.adapter(T::class.java)?.serializeNulls()?.indent(if (prettyPrint) "    " else "")?.toJson(obj)
        ?: throw JsonDataException("Is there an adapter missing for ${T::class.simpleName}?")

fun <T> Moshi.toJson(obj: T, cls: Class<T>, prettyPrint: Boolean = false): String =
    this.adapter(cls)?.serializeNulls()?.indent(if (prettyPrint) "    " else "")?.toJson(obj)
        ?: throw JsonDataException("Is there an adapter missing for ${cls.simpleName}?")

enum class Season {
    SPRING,
    FALL
}

data class CurrentSemester(val year: String, val season: Season)

data class Config(
    val currentSemester: CurrentSemester,
    val attendanceDiscordURL: String,
    val attendanceFolder: String,
    val minutesFolder: String,
    val debugMode: Boolean
)

fun String.titlecase() = this.first().uppercase() + this.substring(1).lowercase()
