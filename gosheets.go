package gosheets

import (
	"context"
	"fmt"
	"os"
	"strings"

	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

type GoogleSheetsClient struct {
	spreadsheetID string
	service       *sheets.Service
}

// NewGoogleSheetsClient initializes a Google Sheets client with the provided
// credentials and spreadsheet ID. It uses a Google Developers service account
// JSON key file to authenticate the requests. You can create a service account
// for your project at https://console.developers.google.com and download a JSON key file.
//
// Parameters:
//   - credentials: The path to the JSON credentials file for authentication.
//   - spreadsheetID: The ID of the Google Sheets spreadsheet to interact with.
//
// Returns:
//   - A pointer to a GoogleSheetsClient instance representing the initialized client.
//   - An error if there was a problem initializing the client, nil otherwise.
func NewGoogleSheetsClient(credentials string, spreadsheetID string) (*GoogleSheetsClient, error) {
	creds, err := os.ReadFile(credentials)
	if err != nil {
		return nil, fmt.Errorf("unable to read credentials file: %v", err)
	}

	config, err := google.JWTConfigFromJSON(creds, sheets.SpreadsheetsScope)
	if err != nil {
		return nil, fmt.Errorf("unable to create JWT config: %v", err)
	}

	client := config.Client(context.Background())
	svc, err := sheets.NewService(context.Background(), option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("unable to create Google Sheets service: %v", err)
	}

	return &GoogleSheetsClient{
		spreadsheetID: spreadsheetID,
		service:       svc,
	}, nil
}

// ReadData reads data from a Google Sheets spreadsheet.
//
// Parameters:
//   - spreadsheetID: The ID of the spreadsheet from which to read data.
//   - readRange: The range of cells to read data from (e.g., "Sheet1!A1:B2").
//
// Returns:
//   - A 2D slice representing the read data, or an error if there was a problem.
func (gs *GoogleSheetsClient) ReadData(readRange string) ([][]interface{}, error) {
	resp, err := gs.service.Spreadsheets.Values.Get(gs.spreadsheetID, readRange).Do()
	if err != nil {
		return nil, fmt.Errorf("unable to retrieve data from Google Sheets: %v", err)
	}
	return resp.Values, nil
}

// AddData adds data to a Google Sheets spreadsheet.
//
// Parameters:
//   - sheetName: The name of the sheet where the data will be added.
//   - startCell: The starting cell used to search for existing data and find a "table"
//     within that range where the data will be appended (e.g., "A1").
//   - values: A 2D slice representing the data to be added. Each inner slice
//     represents a row of data, with each element representing a cell value.
//
// Returns:
//   - An error if there was a problem adding the data to the spreadsheet, nil otherwise.
func (gs *GoogleSheetsClient) AddData(sheetName string, startCell string, values [][]interface{}) error {
	valueRange := &sheets.ValueRange{
		Values: values,
	}

	_, err := gs.service.Spreadsheets.Values.Append(gs.spreadsheetID, sheetName+"!"+startCell, valueRange).ValueInputOption("RAW").Do()
	if err != nil {
		return fmt.Errorf("unable to add data to Google Sheets: %v", err)
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

// DeleteRow deletes a row from a Google Sheets spreadsheet based on a filter.
//
// Parameters:
//   - spreadsheetID: The ID of the spreadsheet from which to delete the row.
//   - data: The 2D slice representing the data to search through.
//   - value: The value to filter on.
//   - column: The column in which to apply the filter.
//
// Returns:
//   - An error if there was a problem deleting the row, nil otherwise.
func (gs *GoogleSheetsClient) DeleteRow(data [][]interface{}, value, column string) error {
	rowIndex := findRowNumber(data, value, column)
	if rowIndex == -1 {
		return fmt.Errorf("unable to find the value %v in column %v", value, column)
	}

	requests := []*sheets.Request{
		{
			DeleteDimension: &sheets.DeleteDimensionRequest{
				Range: &sheets.DimensionRange{
					SheetId:    0, // Sheet ID (can be 0 for the first sheet)
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

	_, err := gs.service.Spreadsheets.BatchUpdate(gs.spreadsheetID, batchUpdate).Do()
	if err != nil {
		return fmt.Errorf("unable to delete row from Google Sheets: %v", err)
	}
	return nil
}

// findRowNumber finds the row number containing a specific value in a given column.
//
// Parameters:
//   - data: The 2D slice representing the data to search through.
//   - value: The value to search for.
//   - column: The column letter in which to search for the value.
//
// Returns:
//   - The row number (1-based index) containing the value, or -1 if not found.
func findRowNumber(data [][]interface{}, value string, column string) int {
	for i, row := range data {
		if len(row) < columnIndex(column) {
			continue // Avoid index out of range error if the row does not have enough columns
		}
		if fmt.Sprintf("%v", row[columnIndex(column)]) == value {
			return i + 1 // Row index is 1-based
		}
	}
	return -1
}

// columnIndex converts a column letter (e.g., "A", "B", ...) to its index (0-based).
//
// Parameters:
//   - column: The column letter to convert.
//
// Returns:
//   - The index of the column.
func columnIndex(column string) int {
	column = strings.ToUpper(column)
	index := 0
	for i, c := range column {
		index = index*26 + int(c-'A') + 1
		if i == 0 {
			index--
		}
	}
	return index
}
