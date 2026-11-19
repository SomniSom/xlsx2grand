// Package vor — модель ведомости объёмов работ и её чтение из xlsx.
package vor

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"xlsx2grand/internal/xlsx"
)

// Item — строка ведомости объёмов работ.
type Item struct {
	Num      int     // № п/п
	NumLSR   string  // № в ЛСР (колонка B) — связь со сметой
	Name     string  // наименование работ
	Unit     string  // ед. изм.
	Quantity float64 // количество
	Formula  string  // формула расчёта (колонка H)
	K        float64 // коэффициент (колонка J)
	HasK     bool
	Links    string // ссылки на чертежи (колонка G)
	Header   string // подзаголовок, под которым стоит строка
	Row      int
}

// Section — раздел ведомости.
type Section struct {
	Caption string
	Num     int
	Headers []string
	Items   []*Item
}

// VOR — ведомость целиком.
type VOR struct {
	Title    string
	Subject  string
	Sections []*Section
	Warnings []string
}

var reSection = regexp.MustCompile(`^Раздел\s+(\d+)\.\s*(.+)$`)

// Parse читает xlsx ведомости объёмов работ.
func Parse(path string) (*VOR, error) {
	book, err := xlsx.Open(path)
	if err != nil {
		return nil, err
	}
	sh := book.Sheets[0]
	v := &VOR{}

	// Шапка: строка 2 — заголовок «Ведомость объёмов работ».
	for r := 1; r <= 5 && r <= sh.MaxRow; r++ {
		if t := sh.Text(1, r); strings.HasPrefix(t, "Ведомость") {
			v.Title = t
		}
	}

	// Колонки: A №, B № в ЛСР, C наименование, D ед., E кол-во, F/G чертежи, H формула, J табл.
	const (
		colNum  = 1
		colLSR  = 2
		colName = 3
		colUnit = 4
		colQty  = 5
		colLink = 7
		colForm = 8
		colK    = 10
	)

	var cur *Section
	var curHeader string
	// Данные начинаются после шапки (строки 4–5 заголовков).
	start := 6
	for r := 1; r <= 10 && r <= sh.MaxRow; r++ {
		if sh.Text(colNum, r) == "№ п/п" {
			// найти строку с числами 1,2,3...
			for rr := r; rr <= r+4; rr++ {
				if sh.Text(colNum, rr) == "1" && sh.Text(colLSR, rr) == "2" {
					start = rr + 1
					break
				}
			}
			break
		}
	}

	for r := start; r <= sh.MaxRow; r++ {
		a := sh.Text(colNum, r)
		b := sh.Text(colLSR, r)
		c := sh.Text(colName, r)

		if a == "" && b == "" && c == "" {
			continue
		}
		// Подписи в конце документа
		if strings.HasPrefix(b, "Составил") || strings.HasPrefix(b, "Проверил") ||
			strings.HasPrefix(c, "[должность") {
			break
		}

		if m := reSection.FindStringSubmatch(a); m != nil {
			n, _ := strconv.Atoi(m[1])
			cur = &Section{Caption: m[2], Num: n}
			curHeader = ""
			v.Sections = append(v.Sections, cur)
			continue
		}

		// Строка позиции: в A номер, в B номер в ЛСР.
		if b != "" && isNum(b) && c != "" {
			n, err := strconv.Atoi(a)
			if err != nil {
				n = 0
			}
			item := &Item{
				Num:     n,
				NumLSR:  b,
				Name:    c,
				Unit:    sh.Text(colUnit, r),
				Formula: strings.TrimSpace(sh.Text(colForm, r)),
				Links:   sh.Text(colLink, r),
				Header:  curHeader,
				Row:     r,
			}
			item.Quantity, _ = sh.Num(colQty, r)
			if k, ok := sh.Num(colK, r); ok {
				item.K = k
				item.HasK = true
			}
			if n == 0 && a != "" {
				v.Warnings = append(v.Warnings, fmt.Sprintf("строка %d: номер в A не число (%q)", r, a))
			}
			if cur == nil {
				cur = &Section{Caption: "Раздел 1", Num: 1}
				v.Sections = append(v.Sections, cur)
			}
			cur.Items = append(cur.Items, item)
			continue
		}

		// Заголовок подраздела: текст в A, остальное пусто.
		if a != "" && b == "" && c == "" && cur != nil {
			cur.Headers = append(cur.Headers, a)
			curHeader = a
		}
	}

	if len(v.Sections) == 0 {
		return nil, fmt.Errorf("%s: не найдено ни одного раздела", path)
	}
	total := 0
	for _, s := range v.Sections {
		total += len(s.Items)
	}
	if total == 0 {
		return nil, fmt.Errorf("%s: не найдено ни одной строки работ", path)
	}
	return v, nil
}

func isNum(s string) bool {
	_, err := strconv.Atoi(strings.TrimSpace(s))
	return err == nil
}
