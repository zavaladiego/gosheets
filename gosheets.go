package gosheets

import (
	"context"
	"fmt"
	"strings"

	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

// GoogleSheetsClient represents a client for interacting with Google Sheets.
//
//   - The service field is used to interact with the Google Sheets API.
//   - The spreadsheetID field is used to store the ID of the Google Sheets spreadsheet to interact with. You can find this ID in the URL of the spreadsheet. For example, the spreadsheet ID in the URL https://docs.google.com/spreadsheets/d/abc1234567/edit#gid=0 is "abc1234567".
//   - The sheetName field is used to store the name of the sheet to interact with in the Google Sheets spreadsheet.
type GoogleSheetsClient struct {
	service       *sheets.Service
	spreadsheetID string
	sheetName     string
}

// NewGoogleSheetsClient initializes a Google Sheets client with the provided
// credentials and spreadsheet ID. It uses a Google Developers service account
// JSON key file to authenticate the requests. You can create a service account
// for your project at https://console.developers.google.com and download a JSON key file.
//
// Parameters:
//   - credentials: The path to the JSON credentials file for authentication.
//
// Returns:
//   - A pointer to a GoogleSheetsClient instance representing the initialized client.
//   - An error if there was a problem initializing the client, nil otherwise.
func NewGoogleSheetsClient(credentials []byte) (*GoogleSheetsClient, error) {
	config, err := google.JWTConfigFromJSON(credentials, sheets.SpreadsheetsScope)
	if err != nil {
		return nil, fmt.Errorf("unable to create JWT config: %v", err)
	}

	client := config.Client(context.Background())
	svc, err := sheets.NewService(context.Background(), option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("unable to create Google Sheets service: %v", err)
	}

	return &GoogleSheetsClient{
		service: svc,
	}, nil
}

// SetSpreadsheetID sets the spreadsheet ID in the GoogleSheetsClient struct.
//
// Parameters:
//   - The ID of the Google Sheets spreadsheet to interact with.
func (gs *GoogleSheetsClient) SetSpreadsheetID(spreadsheetID string) {
	gs.spreadsheetID = spreadsheetID
}

// SetSheetName sets the sheet name in the GoogleSheetsClient struct.
//
// Parameters:
//   - The name of the sheet to interact with in the Google Sheets spreadsheet.
func (gs *GoogleSheetsClient) SetSheetName(sheetName string) {
	gs.sheetName = sheetName
}

// getSheetID retrieves the sheet ID of current sheet set in the GoogleSheetsClient struct.
//
// Returns:
//   - The ID of the sheet, or an error if the sheet was not found.
func (gs *GoogleSheetsClient) getSheetID() (int64, error) {
	err := validateClientFields(gs)
	if err != nil {
		return -1, err
	}

	spreadsheet, err := gs.service.Spreadsheets.Get(gs.spreadsheetID).Do()
	if err != nil {
		return -1, fmt.Errorf("unable to retrieve spreadsheet: %v", err)
	}

	for _, sheet := range spreadsheet.Sheets {
		if sheet.Properties.Title == gs.sheetName {
			return sheet.Properties.SheetId, nil
		}
	}

	return -1, fmt.Errorf("sheet with name %s not found", gs.sheetName)
}

// ReadData reads data from the current set sheet in the GoogleSheetsClient struct.
//
// Parameters:
//   - readRange: The range of cells to read data from (e.g., "A1:B2").
//
// Returns:
//   - A 2D slice representing the read data, or an error if there was a problem.
func (gs *GoogleSheetsClient) ReadData(readRange string) ([][]interface{}, error) {
	err := validateClientFields(gs)
	if err != nil {
		return nil, err
	}

	readRange = gs.sheetName + "!" + readRange

	resp, err := gs.service.Spreadsheets.Values.Get(gs.spreadsheetID, readRange).Do()
	if err != nil {
		return nil, fmt.Errorf("unable to retrieve data from Google Sheets: %v", err)
	}
	return resp.Values, nil
}

// AppendData appends data to the end of the current set sheet in the GoogleSheetsClient struct.
// Parameters:
//   - data: A 2D slice representing the data to be added. Each inner slice
//     represents a row of data, with each element representing a cell value.
//   - range_: The cell used to search for existing data and find a "table"
//     within that range where the data will be appended (e.g., "A1").
//
// Returns:
//   - An error if there was a problem adding the data to the spreadsheet, nil otherwise.
func (gs *GoogleSheetsClient) AppendData(data [][]interface{}, range_ string) error {
	err := validateClientFields(gs)
	if err != nil {
		return err
	}

	valueRange := &sheets.ValueRange{
		Values: data,
	}

	range_ = gs.sheetName + "!" + range_
	_, err = gs.service.Spreadsheets.Values.Append(gs.spreadsheetID, range_, valueRange).ValueInputOption("RAW").Do()
	if err != nil {
		return fmt.Errorf("unable to add data to Google Sheets: %v", err)
	}
	return nil
}

// InsertRowsAtPosition inserts a specified number of rows after a specified row in current set sheet in the GoogleSheetsClient struct. The position is 1-based, with 1 being the header row. Note: Is not possible to insert rows before the header row, giving a position of 0 will result in an error.
//
// Parameters:
//   - data: The data to insert into the spreadsheet.
//   - position: The index of the row after which the new rows will be inserted.
//
// Returns:
//   - An error if there was a problem inserting the rows, nil otherwise.
func (gs *GoogleSheetsClient) InsertRowsAfterPosition(data [][]interface{}, position int64) error {
	sheetID, err := gs.getSheetID()
	if err != nil {
		return fmt.Errorf("unable to retrieve sheet ID: %v", err)
	}

	numRows := int64(len(data))

	insertRequest := &sheets.Request{
		InsertDimension: &sheets.InsertDimensionRequest{
			Range: &sheets.DimensionRange{
				SheetId:    sheetID,
				Dimension:  "ROWS",
				StartIndex: position, // 0 is not a valid index in InsertDimensionRequest
				EndIndex:   position + numRows,
			},
			InheritFromBefore: false,
		},
	}

	batchUpdateRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{insertRequest},
	}

	_, err = gs.service.Spreadsheets.BatchUpdate(gs.spreadsheetID, batchUpdateRequest).Do()
	if err != nil {
		return fmt.Errorf("unable to insert rows at position: %v", err)
	}

	err = gs.AppendData(data, "A"+fmt.Sprint(position))
	if err != nil {
		return fmt.Errorf("unable to append data to Google Sheets: %v", err)
	}

	return nil
}

// InsertRowsAtBeginning inserts a specified number of rows at the beginning of current set sheet in the GoogleSheetsClient struct.The beginning of the sheet is considered to be the row after the header row. Note: This function assumes that the header row is the first row in the sheet.
//
// Parameters:
//   - data: The data to insert into the spreadsheet.
//
// Returns:
//   - An error if there was a problem inserting the rows, nil otherwise.
func (gs *GoogleSheetsClient) InsertRowsAtBeginning(data [][]interface{}) error {
	err := gs.InsertRowsAfterPosition(data, 1)
	if err != nil {
		return fmt.Errorf("unable to insert rows at beginning: %v", err)
	}

	return nil
}

// DeleteRow deletes the first occurrence of a row from a current set sheet in the GoogleSheetsClient struct. Note: This function assumes that the header row is the first row in the sheet.
//
// Parameters:
//   - data: The 2D slice representing the data to search through. Use the ReadData method to get this data.
//   - column: The column in which to apply the filter.
//   - value: The value to filter on.
//
// Returns:
//   - An error if there was a problem deleting the row, nil otherwise.
func (gs *GoogleSheetsClient) DeleteRow(data [][]interface{}, column, value string) error {
	rowIndex := findRowNumber(data, column, value)
	if rowIndex == -1 {
		return fmt.Errorf("unable to find the value %v in column %v", value, column)
	}

	sheetID, err := gs.getSheetID()
	if err != nil {
		return fmt.Errorf("unable to retrieve sheet ID: %v", err)
	}

	requests := []*sheets.Request{
		{
			DeleteDimension: &sheets.DeleteDimensionRequest{
				Range: &sheets.DimensionRange{
					SheetId:    sheetID, // Sheet ID (can be 0 for the first sheet)
					Dimension:  "ROWS",
					StartIndex: int64(rowIndex - 1), // Index of the row to delete (subtract 1 since rows are 0-based)
					EndIndex:   int64(rowIndex),     // Index of the next row
				},
			},
		},
	}

	batchUpdate := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}

	_, err = gs.service.Spreadsheets.BatchUpdate(gs.spreadsheetID, batchUpdate).Do()
	if err != nil {
		return fmt.Errorf("unable to delete row from Google Sheets: %v", err)
	}
	return nil
}

// DataToString converts a 2D slice of interface{} values to a string.
//
// Parameters:
//   - data: The 2D slice of interface{} values to convert.
//
// Returns:
//   - A string representation of the data.
func DataToString(data [][]interface{}) string {
	var result strings.Builder

	for _, row := range data {
		for _, cell := range row {
			result.WriteString(fmt.Sprintf("%v\t", cell))
		}
		result.WriteString("\n")
	}

	return result.String()
}

// findRowNumber finds the row number containing a specific value in a given column.
//
// Parameters:
//   - data: The 2D slice representing the data to search through. Use the ReadData method to get this data.
//   - column: The column letter in which to search for the value.
//   - value: The value to search for.
//
// Returns:
//   - The row number (1-based index) containing the value, or -1 if not found.
func findRowNumber(data [][]interface{}, column, value string) int {
	for i, row := range data {
		if len(row) < columnIndex(column)+1 {
			continue // Avoid index out of range error if the row does not have enough columns
		}
		if fmt.Sprintf("%v", row[columnIndex(column)]) == value {
			return i + 1 // 1-based row index
		}
	}
	return -1
}

// columnIndex converts a column letter (e.g., "A", "Z", "AA", "BX" ... "ZZZ") to its index (0-based).
//
// Parameters:
//   - column: The column letter to convert.
//
// Returns:
//   - The index of the column.
func columnIndex(column string) int {
	column = strings.ToUpper(column)
	index := 0
	for _, c := range column {
		index = index*26 + int(c-'A') + 1
	}
	return index - 1
}

func validateClientFields(gs *GoogleSheetsClient) error {
	if gs.spreadsheetID == "" {
		return fmt.Errorf("spreadsheet ID not set")
	}

	if gs.sheetName == "" {
		return fmt.Errorf("sheet name not set")
	}

	return nil
}
