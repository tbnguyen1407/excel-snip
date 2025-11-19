# excel-snip

Tool to automate screenshots of website URLs in Excel workbook

## Table of Contents

- [Develop](#develop)
- [Run](#run)

## Develop

### Pre-requisites

- Go 1.25+

### Build

```shell
go build .
```

## Run

```shell
// (optional) open browser to perform any login
excel-snip --browse

// execute
excel-snip --book samples/book.xlsx

// help page
excel-snip --help
```
