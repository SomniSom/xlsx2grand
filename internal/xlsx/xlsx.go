// Package xlsx — минимальный читатель .xlsx без внешних зависимостей.
package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"regexp"
	"strconv"
	"strings"
)

// Cell — ячейка листа в A1-адресации.
type Cell struct {
	Col     int     // 1-based
	Row     int     // 1-based
	Value   string  // текстовое представление
	Number  float64 // значение, если ячейка числовая
	IsNum   bool
	Formula string // содержимое <f>, если есть
}

// Sheet — лист с доступом по адресу.
type Sheet struct {
	Name   string
	Rows   map[int][]Cell // строка -> ячейки
	MaxRow int
}

// Book — книга.
type Book struct {
	Sheets []*Sheet
}

// Get возвращает ячейку по колонке и строке (1-based).
func (s *Sheet) Get(col, row int) Cell {
	for _, c := range s.Rows[row] {
		if c.Col == col {
			return c
		}
	}
	return Cell{Col: col, Row: row}
}

// Text возвращает текст ячейки без ведущих и хвостовых пробелов.
func (s *Sheet) Text(col, row int) string {
	return strings.TrimSpace(s.Get(col, row).Value)
}

// Num возвращает число ячейки; ok=false, если её нет или это не число.
func (s *Sheet) Num(col, row int) (float64, bool) {
	c := s.Get(col, row)
	return c.Number, c.IsNum
}

type sharedStrings struct {
	Items []string
}

func (ss *sharedStrings) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "si" {
				var sb strings.Builder
				if err := collectText(d, t, &sb); err != nil {
					return err
				}
				ss.Items = append(ss.Items, sb.String())
			}
		case xml.EndElement:
			if t.Name == start.Name {
				return nil
			}
		}
	}
}

// collectText собирает весь текст внутри элемента, учитывая rich text (<r><t>).
func collectText(d *xml.Decoder, start xml.StartElement, sb *strings.Builder) error {
	depth := 1
	for depth > 0 {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "t" {
				var s string
				if err := d.DecodeElement(&s, &t); err != nil {
					return err
				}
				sb.WriteString(s)
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}
	return nil
}

var cellRef = regexp.MustCompile(`^([A-Z]+)(\d+)$`)

// colIndex переводит "A" -> 1, "AA" -> 27.
func colIndex(s string) int {
	n := 0
	for _, r := range s {
		n = n*26 + int(r-'A'+1)
	}
	return n
}

type rawCell struct {
	R  string `xml:"r,attr"`
	T  string `xml:"t,attr"`
	V  string `xml:"v"`
	F  string `xml:"f"`
	IS *struct {
		Inner []byte `xml:",innerxml"`
	} `xml:"is"`
}

type rawRow struct {
	R     int       `xml:"r,attr"`
	Cells []rawCell `xml:"c"`
}

type rawSheet struct {
	Rows []rawRow `xml:"sheetData>row"`
}

// Open читает книгу целиком.
func Open(path string) (*Book, error) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("открыть xlsx: %w", err)
	}
	defer zr.Close()

	files := map[string]*zip.File{}
	for _, f := range zr.File {
		files[f.Name] = f
	}

	var ss sharedStrings
	if f, ok := files["xl/sharedStrings.xml"]; ok {
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		defer rc.Close()
		data, err := io.ReadAll(rc)
		if err != nil {
			return nil, err
		}
		if err := xml.Unmarshal(data, &ss); err != nil {
			return nil, fmt.Errorf("sharedStrings: %w", err)
		}
	}

	book := &Book{}
	// Порядок листов: xl/worksheets/sheetN.xml (в норме совпадает с порядком в книге).
	for n := 1; ; n++ {
		name := fmt.Sprintf("xl/worksheets/sheet%d.xml", n)
		f, ok := files[name]
		if !ok {
			break
		}
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, err
		}
		var rs rawSheet
		if err := xml.Unmarshal(data, &rs); err != nil {
			return nil, fmt.Errorf("%s: %w", name, err)
		}
		sh := &Sheet{Name: fmt.Sprintf("sheet%d", n), Rows: map[int][]Cell{}}
		for _, rr := range rs.Rows {
			row := rr.R
			if row == 0 {
				continue
			}
			for _, rc := range rr.Cells {
				m := cellRef.FindStringSubmatch(rc.R)
				if m == nil {
					continue
				}
				col := colIndex(m[1])
				rw, _ := strconv.Atoi(m[2])
				cell := Cell{Col: col, Row: rw, Formula: rc.F}
				switch rc.T {
				case "s":
					i, err := strconv.Atoi(strings.TrimSpace(rc.V))
					if err == nil && i >= 0 && i < len(ss.Items) {
						cell.Value = ss.Items[i]
					}
				case "inlineStr":
					if rc.IS != nil {
						var s string
						if err := xml.Unmarshal(rc.IS.Inner, &s); err == nil {
							cell.Value = s
						}
					}
				case "str", "e":
					cell.Value = rc.V
				default:
					cell.Value = rc.V
					if fv, err := strconv.ParseFloat(strings.TrimSpace(rc.V), 64); err == nil {
						cell.Number = fv
						cell.IsNum = true
					}
				}
				sh.Rows[rw] = append(sh.Rows[rw], cell)
				if rw > sh.MaxRow {
					sh.MaxRow = rw
				}
			}
		}
		book.Sheets = append(book.Sheets, sh)
	}
	if len(book.Sheets) == 0 {
		return nil, fmt.Errorf("в %s нет листов", path)
	}
	return book, nil
}
