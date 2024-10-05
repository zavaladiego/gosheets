# GoSheets

GoSheets is a simple Go package wrapper that provides basic functionalities to interact with Google Sheets spreadsheets programmatically. It allows you to read data from, add data to, and delete rows from a Google Sheets spreadsheet.

## Usage Examples

1. **Initialize Google Sheets Service:**

    ```go
    credentials, err := os.ReadFile("path/to/credentials.json")
    gs, err := gosheets.NewGoogleSheetsClient(credentials)
    ```

2. **Set the Spreadsheet ID and sheet name:**

    ```go
    gs.SetSpreadsheetID("spreadsheetID")
    gs.SetSheetName("Sheet1")
    ```

3. **Read Data from current sheet set:**

    ```go
    data, err := gs.ReadData("A:F")
    ```

4. **Append Data to current sheet set:**

    ```go
    values := [][]interface{}{
        {"Value1", "Value2"},
        {"Value3", "Value4"},
    }
    err := gs.AppendData(values, "A1")
    ```

5. **Insert data after a specific row in the current sheet set:**

    ```go
    values := []interface{}{"Value1", "Value2"}
    err := gs.InsertRowsAfterPosition(values, "3")
    ```

6. **Insert data at the beginning of the current sheet set:**

    ```go
    values := []interface{}{"Value1", "Value2"}
    err := gs.InsertRowsAtBeginning(values)
    ```

7. **Delete Row from current sheet set:**

    ```go
    err := gs.DeleteRow(data, "A", value)
    ```

8. **Print Data as string:**

    ```go
    data, err := gs.ReadData("A:F")
    fmt.Println(gs.PrintData(data))
    ```

## Installation

```bash
go get github.com/zavaladiego/gosheets@latest
```

## Requirements

To use this package, ensure you have the following:

- Google Sheets API enabled in your Google Cloud Console. You can enable it by visiting the [Google Cloud Console](https://console.developers.google.com), creating a new project and enabling the [Google Sheets API](https://console.cloud.google.com/marketplace/product/google/sheets.googleapis.com).
- Service account key file (JSON) obtained from the Google Cloud Console. This key file is used for authentication when accessing Google Sheets.
- The google service account must have access to the spreadsheet you want to work with. Ensure that this Google account has been granted the necessary permissions (e.g., edit, read) for the spreadsheet.

## Coverage results

|ok|  github.com/zavaladiego/gosheets	
|--|--|
|execution time| 5.408s|
|coverage| 95.5% of statements|

## Contributing

Contributions are welcome! Feel free to open issues or pull requests.
