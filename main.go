package main

import (
	"errors"
	"flag"
	"fmt"
	"log/slog"
	"os"
	"path"
	"regexp"
	"runtime"
	"strconv"
	"strings"
	"time"

	"github.com/go-rod/rod"
	"github.com/go-rod/rod/lib/devices"
	"github.com/go-rod/rod/lib/launcher"
	"github.com/go-rod/rod/lib/proto"
	"github.com/go-rod/rod/lib/utils"
	"github.com/xuri/excelize/v2"
)

var (
	version = "unset" // set during build with -ldflags
)

func main() {
	// define flags
	var bookPath string
	var sheetNumber int
	var column string
	var filter string
	var template string
	var resolution string
	var timeout int
	var outDirPath string
	var browse bool
	var debug bool
	var versionFlag bool

	// parse flags
	flag.StringVar(&bookPath, "book", "REQUIRED", "Path to XLSX workbook")
	flag.IntVar(&sheetNumber, "sheet", 0, "Worksheet index")
	flag.StringVar(&column, "column", "K", "Column name")
	flag.StringVar(&filter, "filter", "GPLS-", "Cell content filter")
	flag.StringVar(&template, "template", "https://sgtechstack.atlassian.net/browse/__VALUE__", "URL template (use __VALUE__ placeholder for cell content)")
	flag.StringVar(&resolution, "resolution", "1920,1080", "Browser page resolution")
	flag.IntVar(&timeout, "timeout", 60_000, "Browser page timeout (ms)")
	flag.StringVar(&outDirPath, "out", "out", "Path to output directory")
	flag.BoolVar(&debug, "debug", false, "Show browser window during execution")
	flag.BoolVar(&browse, "browse", false, "Open browser")
	flag.BoolVar(&versionFlag, "version", false, "Print version")
	flag.Parse()

	// print version
	if versionFlag {
		fmt.Println(version)
		return
	}

	// open browser
	if browse {
		slog.Info("open browser, Ctrl+C to quit")

		// launch headed browwer
		debugURL := launcher.New().Headless(false).UserDataDir(getUserDataDir()).Delete("enable-automation").MustLaunch()
		browser := rod.New().ControlURL(debugURL).DefaultDevice(devices.Clear).MustConnect()
		defer browser.MustClose()

		browser.Page(proto.TargetCreateTarget{URL: ""})

		// wait for user to interrupt
		for true {
			time.Sleep(5000)
		}
	}

	// validate flags
	var pageW, pageH int
	var e error
	resolutionSplit := strings.Split(resolution, ",")
	if len(resolutionSplit) != 2 {
		exitOnError(errors.New("invalid resolution"))
	}
	pageW, e = strconv.Atoi(resolutionSplit[0])
	exitOnError(e)
	pageH, e = strconv.Atoi(resolutionSplit[1])
	exitOnError(e)

	// build filter
	filterRegex := regexp.MustCompile(filter)

	// load book
	bookFile, e := excelize.OpenFile(bookPath)
	exitOnError(e)
	defer bookFile.Close()

	// get sheet name
	sheetName := bookFile.GetSheetName(sheetNumber)
	if sheetName == "" {
		exitOnError(errors.New("sheet not found"))
	}

	// extract issues
	rows, e := bookFile.GetRows(sheetName)
	exitOnError(e)

	issues := map[string]struct{}{}
	for row := 1; row <= len(rows); row++ {
		// get cell content
		cellAddr := column + strconv.Itoa(row)
		cellValue, e := bookFile.GetCellValue(sheetName, cellAddr)
		if e != nil {
			slog.Warn("skipping cell", "addr", cellAddr, "error", e)
			continue
		}
		cellValue = strings.TrimSpace(cellValue)

		// skip unmatch
		if !filterRegex.MatchString(cellValue) {
			continue
		}

		issues[cellValue] = struct{}{}
	}

	if len(issues) == 0 {
		return
	}

	// ensure output
	e = os.MkdirAll(outDirPath, 0755)
	exitOnError(e)

	// create browser
	debugURL := launcher.New().Headless(!debug).UserDataDir(getUserDataDir()).MustLaunch()
	browser := rod.New().ControlURL(debugURL).MustConnect()
	defer browser.MustClose()
	page := browser.MustPage()
	defer page.MustClose()
	e = page.SetViewport(&proto.EmulationSetDeviceMetricsOverride{Width: pageW, Height: pageH})
	exitOnError(e)

	// take snapshots
	for issue := range issues {
		// construct url
		url := strings.ReplaceAll(template, "__VALUE__", issue)
		// todo: sanitize filename
		outFileName := issue + ".png"
		outFilePath := path.Join(outDirPath, outFileName)

		slog.Info("match", "in", issue, "url", url, "out", outFilePath)

		// navigate
		pageWithTimeout := page.Timeout(time.Duration(timeout) * time.Millisecond)
		e = pageWithTimeout.WaitStable(1 * time.Second)
		e = pageWithTimeout.Navigate(url)
		if e != nil {
			slog.Warn("skipping screenshot", "url", url, "error", e)
			continue
		}

		if e != nil {
			slog.Warn("skipping screenshot", "url", url, "error", e)
			continue
		}

		// capture
		img, e := pageWithTimeout.Screenshot(false, nil)
		if e != nil {
			slog.Warn("skipping screenshot", "url", url, "error", e)
			continue
		}

		// persist
		e = utils.OutputFile(outFilePath, img)
		if e != nil {
			slog.Warn("skipping screenshot", "url", url, "error", e)
			continue
		}
	}
}

func getUserDataDir() string {
	cacheDir := ""
	switch runtime.GOOS {
	case "windows":
		cacheDir = os.Getenv("LOCALAPPDATA")
	default:
		xdg_cache_dir := os.Getenv("XDG_RUNTIME_DIR")
		if xdg_cache_dir != "" {
			cacheDir = xdg_cache_dir
		} else {
			cacheDir = path.Join(os.Getenv("HOME"), ".cache")
		}
	}
	return path.Join(cacheDir, "excel-fill", "userdata")
}
func exitOnError(e error) {
	if e != nil {
		slog.Error("terminating", "error", e)
		os.Exit(1)
	}
}
