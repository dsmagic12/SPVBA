# SPVBA
A library of VBA code designed to work with data in a SharePoint list/library

## Setup
1) Create an MS Access database and import SharePoint list(s) as linked tables
2) For each list you wish to work with, do the following
   A) Copy the table object
   B) Paste the table into the same database window
   C) Give the new object a name of `"archive" + whatever the original linked table's name was`
   D) Choose the option to paste the object as `Structure Only (local table)`
   E) Open the new table in Design View
   F) Add a `Number` column named `ArchiveID`
   G) Add a `Date/Time` column named `ArchiveDateTime` and set its `Default Value` property to `=Now()`
   H) Save your changes and close the table

## MoveDataToSharePoint
1) Update the code to use your desired `SourceTableName` and `DestinationTableName`
2) Update the code's query of the `SourceTableName` to only include coumns/fields you wish to migrate data for --- omit any fields that are not editable in the `DestinationTableName`
3) Run the code to loop through the rows/records in `SourceTableName`, migrating their data to the `DestinationTableName` field-by-field

## ArchiveDataFromSharePoint
1) Update the code to handle archiving of attached files:
   A) `Const ATTACHMENTS_PATH As String = "{full_path_of_folder_to_download_attached_files_to}"`
2) Create the `AttachmentPaths` table:
   A) Add a field named `ID` with type `AutoNumber` and make it the Primary Key
   B) Add a field named `ArchiveID` with type `Number`
   C) Add a field named `ArchiveDateTime` with type `Date/Time` and set its `Default Value` property to `=Now()`
3) Update the code to use your desired `SourceTableName` and `DestinationTableName`
   A) `SourceTableName` should be the name of the linked table that contains the SharePoint list items you wish to archive
   B) `DestinationTableName` should be the name of the local table you created during the Setup step at the beginning of this readme
4) Update the code's query of the `SourceTableName` to only include coumns/fields you wish to archive data from --- omit any fields that you don't care about
5) Run the code to loop through the rows/records in `SourceTableName`, copying the item column values into the `DestinationTableName` field-by-field
6) Files attached to the source items will be downloaded as `({itemId}) {Attached_File_FileNameAndExtension}` to the path specified at the constant `ATTACHMENTS_PATH` 
   - Note: If the file already exists, it will be overwritten!
