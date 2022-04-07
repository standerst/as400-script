const reader = require('xlsx')

const file = reader.readFile(`${__dirname}/_AS400_REPORT(AutoRecovered).xlsx`)
const sheet = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[0]])

let data = []

sheet.forEach((res) => {
    let code = res.SUMMARY.split('-')[3]
    res.Message_code = code

    data.push(res)
})

var cleaned_sheet = data.filter((data, index, self) =>
    index == self.findIndex((d) => (d.NODE === data.NODE && d.SUMMARY.split('-')[5] === data.SUMMARY.split('-')[5])))

const ws = reader.utils.json_to_sheet(cleaned_sheet)
  
reader.utils.book_append_sheet(file, ws, "Result Sheet")

reader.writeFile(file, `${__dirname}/AS400_REPORT.xlsx`)