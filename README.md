# ICON

## Importing a profiler workbook

From the home page, download the generated profiler template, enter values in
the year columns, and upload the completed `.xlsx` file. Keep the profiler
variable labels unchanged. The app detects the analysis period from the first
and last years containing values, loads those inputs into the profiler, and
autosaves them to the Complete input profile. Enter a field/project name at the
top of the profiler; the name is saved in the database workbook's
`Project Information` sheet and retained on autosave. The database is saved as
a sanitized `database_<project>_<profile>_<id>.xlsx` filename. Changing the
project name renames the active workbook without overwriting another file.
