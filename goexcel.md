
```
	err := goexcel.ExcelRun("edinburgh.xlsx", "newyork", func(f *goexcel.ExcelHandle) error {
		row := f.GetRowNumber()
		if row == 0 {
			f.AppendRowCell(1, 2, 3, 4, 5)
		}
		f.AppendRowCell(4, 5, 3, 3, 3, 6)
		fmt.Println(f.GetAllRows())
		a := f.GetAllRows()
		b := a[len(a)-1][0]
		fmt.Println(">>>", b)
		return nil
	})

	if err != nil {
		fmt.Println(err)
	}
```