package alpha.sig

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
