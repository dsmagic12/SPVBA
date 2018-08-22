# SPVBA
A library of VBA code designed to work with data in a SharePoint list/library

## MoveDataToSharePoint
1) Update the code to use your desired `SourceTableName` and `DestinationTableName`
2) Update the code's query of the `SourceTableName` to only include coumns/fields you wish to migrate data for --- omit any fields that are not editable in the `DestinationTableName`
3) Run the code to loop through the rows/records in `SourceTableName`, migrating their data to the `DestinationTableName` field-by-field
