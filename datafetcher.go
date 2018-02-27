package hoboexcel

// Fetch next row, if no more row exists return nil
type RowFetcher interface {
	NextRow() []string
}

// Fetch next sheet, if no more row exists return nil
type SheetFetcher interface {
	NextSheet() XlsxSheet
}
