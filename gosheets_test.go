package gosheets

import (
	"os"
	"testing"
)

var client *GoogleSheetsClient

func TestMain(m *testing.M) {
	// Read JSON file
	credentials, err := os.ReadFile("PATH/TO/CREDENTIALS.json")
	if err != nil {
		panic("Error reading credentials file: " + err.Error())
	}

	// Create a new Google Sheets client
	client, err = NewGoogleSheetsClient(credentials)
	if err != nil {
		panic("Error creating Google Sheets client: " + err.Error())
	}

	// Run tests
	code := m.Run()

	// Exit with the code from the tests
	os.Exit(code)
}

func resetClient() {
	client.SetSpreadsheetID("SPREADSHEET_ID")
	client.SetSheetName("Sheet1")
}

func TestNewGoogleSheetsClient(t *testing.T) {
	resetClient()

	// Test cases
	tests := []struct {
		name    string
		creds   []byte
		wantErr bool
	}{
		{
			name:    "Invalid credentials",
			creds:   nil,
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := NewGoogleSheetsClient(tt.creds)
			if (err != nil) != tt.wantErr {
				t.Errorf("NewGoogleSheetsClient() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestGetSheetID(t *testing.T) {
	// Test cases
	tests := []struct {
		name                  string
		sheetName             string
		spreadsheetID         string
		validateSpreadsheetID bool
		wantErr               bool
	}{
		{
			name:      "Valid sheet name",
			sheetName: "Sheet1",
			wantErr:   false,
		},
		{
			name:      "Invalid sheet name",
			sheetName: "Sheet2",
			wantErr:   true,
		},
		{
			name:      "Empty sheet name",
			sheetName: "",
			wantErr:   true,
		},
		{
			name:                  "Empty spreadsheet ID",
			sheetName:             "Sheet1",
			spreadsheetID:         "",
			validateSpreadsheetID: true,
			wantErr:               true,
		},
		{
			name:                  "Invalid spreadsheet ID",
			sheetName:             "Sheet1",
			spreadsheetID:         "123456789",
			validateSpreadsheetID: true,
			wantErr:               true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			resetClient()
			client.SetSheetName(tt.sheetName)

			if tt.validateSpreadsheetID {
				client.SetSpreadsheetID(tt.spreadsheetID)
			}

			_, err := client.getSheetID()
			if (err != nil) != tt.wantErr {
				t.Errorf("getSheetID() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestReadData(t *testing.T) {
	resetClient()

	// Test cases
	tests := []struct {
		name                  string
		sheetName             string
		spreadsheetID         string
		validateSpreadsheetID bool
		readRange             string
		wantErr               bool
	}{
		{
			name:      "Valid read range",
			sheetName: "Sheet1",
			readRange: "A1:B4",
			wantErr:   false,
		},
		{
			name:      "Invalid read range (non-existent sheet)",
			sheetName: "Sheet2",
			readRange: "A1",
			wantErr:   true,
		},
		{
			name:      "Invalid read range (empty sheet name)",
			sheetName: "",
			readRange: "A1",
			wantErr:   true,
		},
		{
			name:                  "Invalida spreadsheet ID",
			spreadsheetID:         "123456789",
			validateSpreadsheetID: true,
			sheetName:             "Sheet1",
			readRange:             "A1",
			wantErr:               true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if tt.validateSpreadsheetID {
				client.SetSpreadsheetID(tt.spreadsheetID)
			}

			client.SetSheetName(tt.sheetName)

			data, err := client.ReadData(tt.readRange)
			if (err != nil) != tt.wantErr {
				t.Errorf("ReadData() error = %v, wantErr %v", err, tt.wantErr)
			} else {
				t.Logf("Read data: %v", data)
			}
		})
	}
}

func TestAppendData(t *testing.T) {
	resetClient()

	// Test cases
	tests := []struct {
		name                  string
		data                  [][]interface{}
		range_                string
		sheetName             string
		spreadsheetID         string
		validateSpreadsheetID bool
		wantErr               bool
	}{
		{
			name:      "Valid data",
			data:      [][]interface{}{{"Value1", "Value2"}, {"Value3", "Value4"}},
			range_:    "A1",
			sheetName: "Sheet1",
			wantErr:   false,
		},
		{
			name:      "Valid data (empty)",
			data:      [][]interface{}{},
			range_:    "A1",
			sheetName: "Sheet1",
			wantErr:   false,
		},
		{
			name:      "Invalid range",
			data:      [][]interface{}{{"Value1", "Value2"}, {"Value3", "Value4"}},
			range_:    "",
			sheetName: "Sheet1",
			wantErr:   true,
		},
		{
			name:      "Empty sheet name",
			data:      [][]interface{}{{"Value1", "Value2"}, {"Value3", "Value4"}},
			range_:    "A1",
			sheetName: "",
			wantErr:   true,
		},
		{
			name:                  "Empty spreadsheet ID",
			data:                  [][]interface{}{{"Value1", "Value2"}, {"Value3", "Value4"}},
			range_:                "A1",
			sheetName:             "Sheet1",
			spreadsheetID:         "",
			validateSpreadsheetID: true,
			wantErr:               true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if tt.validateSpreadsheetID {
				client.SetSpreadsheetID(tt.spreadsheetID)
			}

			client.SetSheetName(tt.sheetName)

			err := client.AppendData(tt.data, tt.range_)
			if (err != nil) != tt.wantErr {
				t.Errorf("AppendData() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestInsertRowsAfterPosition(t *testing.T) {
	resetClient()

	// Test cases
	tests := []struct {
		name                  string
		data                  [][]interface{}
		position              int
		sheetName             string
		spreadsheetID         string
		validateSpreadsheetID bool
		wantErr               bool
	}{
		{
			name:      "Valid position 3",
			data:      [][]interface{}{{"111", "222"}, {"333", "444"}, {"555", "666"}},
			position:  3,
			sheetName: "Sheet1",
			wantErr:   false,
		},
		{
			name:      "Empty sheet name",
			data:      [][]interface{}{{"111", "222"}, {"333", "444"}, {"555", "666"}},
			position:  3,
			sheetName: "",
			wantErr:   true,
		},
		{
			name:                  "Empty spreadsheet ID",
			data:                  [][]interface{}{{"111", "222"}, {"333", "444"}, {"555", "666"}},
			position:              3,
			sheetName:             "Sheet1",
			spreadsheetID:         "",
			validateSpreadsheetID: true,
			wantErr:               true,
		},
		{
			name:      "Invalid position 0",
			data:      [][]interface{}{{"111", "222"}, {"333", "444"}, {"555", "666"}},
			position:  0,
			sheetName: "Sheet1",
			wantErr:   true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if tt.validateSpreadsheetID {
				client.SetSpreadsheetID(tt.spreadsheetID)
			}

			client.SetSheetName(tt.sheetName)

			err := client.InsertRowsAfterPosition(tt.data, int64(tt.position))
			if (err != nil) != tt.wantErr {
				t.Errorf("InsertRowsAfterPosition() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestInsertRowsAtBeginning(t *testing.T) {
	resetClient()

	// Test cases
	tests := []struct {
		name                  string
		data                  [][]interface{}
		sheetName             string
		spreadsheetID         string
		validateSpreadsheetID bool
		wantErr               bool
	}{
		{
			name:      "Valid data",
			sheetName: "Sheet1",
			data:      [][]interface{}{{"555", "666"}, {"777", "888"}, {"999", "000"}},
			wantErr:   false,
		},
		{
			name:      "Empty sheet name",
			data:      [][]interface{}{{"555", "666"}, {"777", "888"}, {"999", "000"}},
			sheetName: "",
			wantErr:   true,
		},
		{
			name:                  "Empty spreadsheet ID",
			data:                  [][]interface{}{{"555", "666"}, {"777", "888"}, {"999", "000"}},
			sheetName:             "Sheet1",
			spreadsheetID:         "",
			validateSpreadsheetID: true,
			wantErr:               true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if tt.validateSpreadsheetID {
				client.SetSpreadsheetID(tt.spreadsheetID)
			}

			client.SetSheetName(tt.sheetName)

			err := client.InsertRowsAtBeginning(tt.data)
			if (err != nil) != tt.wantErr {
				t.Errorf("InsertRowsAtBeginnig() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestDeleteRow(t *testing.T) {
	resetClient()

	// Test cases
	data, err := client.ReadData("A:B")
	if err != nil {
		t.Fatalf("Error reading data from Google Sheets: %v", err)
	}

	tests := []struct {
		name                  string
		data                  [][]interface{}
		column                string
		value                 string
		sheetName             string
		spreadsheetID         string
		validateSpreadsheetID bool
		wantErr               bool
	}{
		{
			name:      "Existing value in given column",
			data:      data,
			column:    "A",
			value:     "Value1", // Make sure this value exists in the data
			sheetName: "Sheet1",
			wantErr:   false,
		},
		{
			name:      "Invalid data (empty)",
			data:      [][]interface{}{},
			column:    "A",
			value:     "Value1",
			sheetName: "Sheet1",
			wantErr:   true,
		},
		{
			name:      "Non-existent value in given column",
			data:      data,
			column:    "A",
			value:     "Value99",
			sheetName: "Sheet1",
			wantErr:   true,
		},
		{
			name:      "Empty sheet name",
			data:      data,
			column:    "A",
			value:     "Value1",
			sheetName: "",
			wantErr:   true,
		},
		{
			name:                  "Empty spreadsheet ID",
			data:                  data,
			column:                "A",
			value:                 "Value1",
			sheetName:             "Sheet1",
			spreadsheetID:         "",
			validateSpreadsheetID: true,
			wantErr:               true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if tt.validateSpreadsheetID {
				client.SetSpreadsheetID(tt.spreadsheetID)
			}

			client.SetSheetName(tt.sheetName)

			err := client.DeleteRow(tt.data, tt.column, tt.value)
			if (err != nil) != tt.wantErr {
				t.Errorf("DeleteRow() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestDataToString(t *testing.T) {
	// Test cases
	tests := []struct {
		name string
		data [][]interface{}
		want string
	}{
		{
			name: "Valid data (same number of columns)",
			data: [][]interface{}{{"Value1", "Value2"}, {"Value3", "Value4"}},
			want: "Value1\tValue2\t\nValue3\tValue4\t\n",
		},
		{
			name: "Valid data different number of columns",
			data: [][]interface{}{{"Value1", "Value2"}, {"Value3"}},
			want: "Value1\tValue2\t\nValue3\t\n",
		},
		{
			name: "Valid data (empty)",
			data: [][]interface{}{},
			want: "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := DataToString(tt.data); got != tt.want {
				t.Errorf("DataToString() = %v, want %v", got, tt.want)
			} else {
				t.Logf("\nData as string:\n%v", got)
			}
		})
	}
}

func TestFindRowNumber(t *testing.T) {
	// Test cases
	tests := []struct {
		name   string
		data   [][]interface{}
		column string
		value  string
		want   int
	}{
		{
			name:   "Existing value in given column",
			data:   [][]interface{}{{"Value1", "Value2"}, {"Value3", "Value4"}},
			column: "A",
			value:  "Value3",
			want:   2,
		},
		{
			name:   "Non-existent value in given column",
			data:   [][]interface{}{{"Value1", "Value2"}, {"Value3", "Value4"}},
			column: "A",
			value:  "Value99",
			want:   -1,
		},
		{
			name:   "Invalid data (empty)",
			data:   [][]interface{}{},
			column: "A",
			value:  "Value1",
			want:   -1,
		},
		{
			name:   "Invalid column",
			data:   [][]interface{}{{"Value1", "Value2"}, {"Value3", "Value4"}},
			column: "C",
			value:  "Value1",
			want:   -1,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := findRowNumber(tt.data, tt.column, tt.value); got != tt.want {
				t.Errorf("FindRowNumber() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestColumnIndex(t *testing.T) {
	// Test cases
	tests := []struct {
		name   string
		column string
		want   int
	}{
		{
			name:   "Valid column",
			column: "A",
			want:   0,
		},
		{
			name:   "Valid column",
			column: "B",
			want:   1,
		},
		{
			name:   "Valid column",
			column: "Z",
			want:   25,
		},
		{
			name:   "Valid column",
			column: "AA",
			want:   26,
		},
		{
			name:   "Valid column",
			column: "BX",
			want:   75,
		},
		{
			name:   "Valid column",
			column: "ZZZ",
			want:   18277,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := columnIndex(tt.column); got != tt.want {
				t.Errorf("ColumnIndex() = %v, want %v", got, tt.want)
			}
		})
	}
}
