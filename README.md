# GoSheets

GoSheets is a simple Go package that provides basic functionalities to interact with Google Sheets spreadsheets programmatically. It allows you to read data from, add data to, and delete rows from a Google Sheets spreadsheet.

## Usage Examples

1. **Initialize Google Sheets Service:**

    ```go
    gosheets.InitGoogleSheetsService("path/to/credentials.json")
    ```

2. **Read Data from Google Sheets:**

    ```go
    data, err := gosheets.ReadData("spreadsheetID", "Sheet1!A1:B2")
    ```

3. **Add Data to Google Sheets:**

    ```go
    values := [][]interface{}{
        {"Value1", "Value2"},
        {"Value3", "Value4"},
    }
    err := gosheets.AddData("spreadsheetID", values)
    ```

4. **Delete Row from Google Sheets:**

    ```go
    err := gosheets.DeleteRow("spreadsheetID", data, "Value", "A")
    ```

## Installation

```bash
go get -u github.com/dzvCode/gosheets
```

## Requirements

To use this package, ensure you have the following:

- Google Sheets API enabled in your Google Cloud Console. You can enable it by visiting the [Google Cloud Console](https://console.developers.google.com), creating a new project and enabling the [Google Sheets API](https://console.cloud.google.com/marketplace/product/google/sheets.googleapis.com).
- Service account key file (JSON) obtained from the Google Cloud Console. This key file is used for authentication when accessing Google Sheets.
- The google service account must have access to the spreadsheet you want to work with. Ensure that this Google account has been granted the necessary permissions (e.g., edit, read) for the spreadsheet.

## Contributing

Contributions are welcome! Feel free to open issues or pull requests.