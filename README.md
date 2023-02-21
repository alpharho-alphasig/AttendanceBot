# Attendance Bot
## Building
- Run `./gradlew shadowJar` on Linux, or `gradlew.bat shadowJar` on Windows.
- Attendance.jar will be built.

## Running
- Run `java -jar Attendance.jar`.

## How to update for the new semester
- At the top of the file, change the `minutesDirectory`.
- If the attendance sheet has changed, change `outputFile`.
- If the attendance sheet has not changed, change `sheetName` for the current semester.

## Troubleshooting
- Make sure `useTestDirectories` is false.
- Make sure the attendance is the FIRST table in the minutes.
- If you're having gradle troubles, run `./gradlew wrapper` or `gradlew.bat wrapper` to update your gradle version.

