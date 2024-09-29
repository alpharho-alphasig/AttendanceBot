package alpha.sig

import com.squareup.moshi.Moshi
import com.squareup.moshi.kotlin.reflect.KotlinJsonAdapterFactory
import jakarta.xml.bind.JAXBElement
import org.docx4j.openpackaging.io.SaveToZipFile
import org.docx4j.openpackaging.packages.OpcPackage
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage
import org.docx4j.openpackaging.packages.WordprocessingMLPackage
import org.docx4j.openpackaging.parts.Part
import org.docx4j.openpackaging.parts.PartName
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart
import org.docx4j.wml.Tbl
import org.docx4j.wml.Tc
import org.docx4j.wml.Tr
import org.xlsx4j.jaxb.Context
import org.xlsx4j.sml.STCellType
import java.io.File
import java.net.URI
import java.net.http.HttpClient
import java.net.http.HttpRequest
import java.net.http.HttpResponse
import java.nio.file.Files
import java.nio.file.Paths
import java.time.LocalDate
import java.time.format.DateTimeFormatter

// Update these!
const val sheetName = "F2022"

typealias RosterNumber = Int

data class Brother(val name: String, val attendance: String)

fun findNewestMinutes(minutesDir: File): File? {
    val files = minutesDir.listFiles { _, name ->
        name.lowercase().startsWith("meeting minutes")
    } ?: return null
    // Convert the files into a map where the key is the date and the value is the file
    val fileMap = files.associateBy {
        var dateStr = it.nameWithoutExtension.lowercase().substringAfter("meeting minutes ")
        if (dateStr.substringBefore("-").toInt() < 10) // Add a leading zero to the day if it's missing
            dateStr = "0$dateStr"
        LocalDate.parse(dateStr, DateTimeFormatter.ofPattern("MM-dd-yy"))
    }

    // Return the newest minutes
    val date = fileMap.keys.max()
    return fileMap[date]
}

fun extractAttendanceFromDoc(newestMinutes: File): Map<RosterNumber, Brother> {
    val wordPackage = WordprocessingMLPackage.load(newestMinutes)
    val mainDocumentPart = wordPackage.mainDocumentPart
    val brothers = sortedMapOf<RosterNumber, Brother>()
    val tbl = mainDocumentPart.content.first {
        it is JAXBElement<*> && it.declaredType.name.endsWith("Tbl")
    }.unwrap<Tbl>()

    for (row in tbl.content) {
        // Pull out three cells at once, looking like [["741"], ["Jonathan Lewis"], ["P"]].
        for (tuple in (row as Tr).content.windowed(3, 3)) {
            val (roster, name, attendance) = tuple.map {
                it.unwrap<Tc>().content[0].toString() // Pull the values out as strings
            }
            brothers[roster.toInt()] = Brother(name, attendance)
        }
    }
    return brothers
}

/**
 * Recursively go through the relationships in the file to find all the worksheet parts.
 */
fun getWorksheetParts(
    wordMLPackage: OpcPackage,
    relationshipsPart: RelationshipsPart = wordMLPackage.relationshipsPart,
    worksheetParts: MutableList<WorksheetPart> = mutableListOf(),
    handled: MutableSet<Part> = mutableSetOf()
): List<WorksheetPart> {
    for (relationship in relationshipsPart.getRelationships().relationship) {
        val part = relationshipsPart.getPart(relationship)

        // Don't try to find relationships between multiple xlsx files
        if ((relationship.targetMode != null && relationship.targetMode.equals("External")) || part in handled) {
            continue
        }

        handled.add(part) // Make sure you don't try to traverse the same parts again to avoid cycles
        if (part.relationshipsPart != null) {
            getWorksheetParts(wordMLPackage, part.relationshipsPart, worksheetParts, handled) // Recurse
        }

        // Add any WorksheetParts we find into the list.
        if (part is WorksheetPart) {
            worksheetParts.add(part)
        }
    }
    return worksheetParts
}

val microsoftEpoch: LocalDate = LocalDate.of(1900, 1, 1)


fun appendToSpreadsheet(brothers: Map<RosterNumber, Brother>, sheetFile: File): SpreadsheetMLPackage {
    val spreadsheetPackage = SpreadsheetMLPackage.load(sheetFile)

    // Grab the shared strings table, put it in list and map format for convenience
    val sharedStrings = spreadsheetPackage.parts.get(PartName("/xl/sharedStrings.xml")) as SharedStrings
    val stringList = sharedStrings.contents.si.map { it.t.value.toString() }
    val stringMap = stringList.mapIndexed { k, v -> v to k }.toMap()

    // Get the sheet and its data
    val worksheetParts = getWorksheetParts(spreadsheetPackage)
    val sheetId = spreadsheetPackage.workbookPart.contents.sheets.sheet.first { it.name == sheetName }.sheetId - 1
    val sheetdata = worksheetParts[sheetId.toInt()].jaxbElement.sheetData

    // Figure out which column corresponds to this week
    val now = LocalDate.now()
    val dates = sheetdata.row.first().c.subList(3).map { microsoftDateToLocalDate(it.v.toLong()) }
    val nextMonday = now.plusDays((8 - now.dayOfWeek.value.toLong()) % 7)
    val columnLetter = "" + ('C' + dates.indexOf(nextMonday))

    // Go through all the rows excluding the first one
    for (row in sheetdata.row.subList(1)) {
        val roster = row.c.first().v.toInt()
        val attendance = brothers[roster]?.attendance ?: continue
        val styleIndex = row.c.last().s
        // Add the new cell for the brother's attendance
        val newCell = Context.getsmlObjectFactory().createCell()
        if (attendance in stringMap) {
            // If it's in the shared strings (which P, A, E, EL, etc. should be already)
            newCell.t = STCellType.S // Cell type (String from SharedStrings)
            newCell.v = stringMap[attendance].toString() // Value, the index of the shared string
            newCell.r = "$columnLetter${row.r}" // Reference, like H7
            newCell.s = styleIndex // Style index, matches the index of whatever is to its left
        } else {
            // If for some reason, it's not in the shared strings, make it inline to save the headache.
            newCell.t = STCellType.INLINE_STR
            newCell.`is` = inlineStr(attendance)
            newCell.r = "$columnLetter${row.r}"
            newCell.s = styleIndex
        }
        row.c.add(newCell)
    }

    return spreadsheetPackage
}

fun getResultsFromSpreadsheet(spreadsheetPackage: SpreadsheetMLPackage): Map<RosterNumber, Pair<String, Float>> {
    // Get the sheet and its data
    val worksheetParts = getWorksheetParts(spreadsheetPackage)
    val sheetId = spreadsheetPackage.workbookPart.contents.sheets.sheet.first { it.name == sheetName }.sheetId - 1
    val sheetdata = worksheetParts[sheetId.toInt()].jaxbElement.sheetData

    return sheetdata.row.associate { row ->
        row.c.first().v.toInt() to Pair(row.c[1].v, row.c[2].v.toFloat())
    }.toSortedMap()
}

fun sendDiscordMessage(msg: String, discordWebhookURL: String, debugMode: Boolean) {
    print(msg)
    if (!debugMode) {
        val response = HttpClient.newHttpClient()
            .send(
                HttpRequest.newBuilder(URI.create(discordWebhookURL))
                    .POST(HttpRequest.BodyPublishers.ofString("{\"content\": \"$msg\"}"))
                    .headers("Content-Type: application/json")
                    .build(), HttpResponse.BodyHandlers.ofString()
            )
        println(response) // For debugging purposes
    }
}

fun loadConfig(moshi: Moshi): Config {
    return moshi.fromJson<Config>(Files.readString(Paths.get("opt", "bots", "config.json")))
}

fun main() {
    val moshi = Moshi.Builder()
        .add(KotlinJsonAdapterFactory())
        .build()
    val config = loadConfig(moshi)
    val useTestDirectories = config.debugMode
    val minutesDirectory = File(
        if (useTestDirectories) "/opt/bots/Attendance/testFiles/"
        else config.minutesFolder
    )
    val outputFile = File(
        if (useTestDirectories) "/opt/bots/Attendance/testFiles/Attendance Fall 2022.xlsx"
        else "${config.attendanceFolder}/${config.currentSemester.season.name.first()}${config.currentSemester.year}/Attendance ${config.currentSemester.season.name.titlecase()} ${config.currentSemester.year}.xlsx"
    )
    val newestMinutes = findNewestMinutes(minutesDirectory)
    if (newestMinutes == null) {
        println("File not found in $minutesDirectory.")
        return
    }
    val brothers = extractAttendanceFromDoc(newestMinutes)
    val sheetPkg = appendToSpreadsheet(brothers, outputFile)
    val results = getResultsFromSpreadsheet(sheetPkg)
    sendDiscordMessage("Attendance report:\n${results.map { (roster, result) ->
        val (name, score) = result
        "$roster, $name: $score${when { score >= 3 -> " (GBS)"; score >= 2 -> " (CBS)"; else -> ""}}"
    }.joinToString("\n")}", config.attendanceDiscordURL, useTestDirectories)
    @Suppress("DEPRECATION")
    SaveToZipFile(sheetPkg).save(
        if (useTestDirectories) File("/opt/bots/Attendance/testFiles/Attendance Fall 2022 after.xlsx")
        else outputFile
    )
}
