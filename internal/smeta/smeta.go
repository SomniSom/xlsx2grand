// Package smeta — модель сметы и чтение выгрузки ЛСР (ГРАНД-Смета) из xlsx.
package smeta

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"xlsx2grand/internal/xlsx"
)

// Resource — ресурс внутри сметной позиции.
type Resource struct {
	Kind     string  // "Tzr", "Mch", "Mat"
	Code     string  // шифр ресурса
	Caption  string  // наименование
	Units    string  // ед. изм.
	Quantity float64 // количество на единицу
	K        float64 // коэффициент (колонка J сметы)
	HasK     bool
	Base     float64 // базовая цена (L)
	HasBase  bool
	Index    float64 // индекс (M)
	HasIndex bool
	Curr     float64 // текущая цена (N)
	HasCurr  bool
	Row      int
}

// Koeff — коэффициент позиции (строка «Прил.1.12 п.3.187_… ОЗП=1,2; ТЗ=1,2»).
type Koeff struct {
	Caption string // наименование коэффициента (до списка «ОЗП=…»)
	Code    string // первая часть наименования до «_» (Прил.1.12 п.3.187)
	OZP     float64
	HasOZP  bool
	EM      float64
	HasEM   bool
	ZPM     float64
	HasZPM  bool
	TZ      float64
	HasTZ   bool
	TZM     float64
	HasTZM  bool
	EMRash  bool // «ЭМ=… к расх.»
	Raw     string
}

// Position — позиция локального сметного расчёта.
type Position struct {
	Number   int
	Code     string
	Caption  string
	Units    string
	Quantity float64
	K        float64
	HasK     bool
	Res      []Resource
	Koeffs   []Koeff
	Row      int
}

// Chapter — раздел.
type Chapter struct {
	Caption   string
	Headers   []string
	Positions []*Position
}

// Estimate — шапка + разделы.
type Estimate struct {
	LocNum        string
	Object        string
	Constr        string
	Description   string
	Basis         string
	Approved      string // нормативная база для DBComment
	CurrPriceDate string
	RegionCode    string // "52"
	RegionName    string
	Composer      string
	Chapters      []*Chapter
	Warnings      []string
}

var (
	reChapter    = regexp.MustCompile(`^Раздел\s+(\d+)\.\s*(.+)$`)
	rePosition   = regexp.MustCompile(`^\d+$`)
	reSimpleCode = regexp.MustCompile(`^\d{1,3}-\d{1,3}$`)
	codePrefixes = []string{"ГЭСН", "ФСБЦ", "ТЦ_", "ФЕР", "ГЭСНм", "ГЭСНр", "02-15", "ФСЭМ"}
)

func isCode(s string) bool {
	for _, p := range codePrefixes {
		if strings.HasPrefix(s, p) {
			return true
		}
	}
	// перевозка вида 02-15-1-01-0015
	if len(s) > 4 && s[2] == '-' && len(s) < 30 && strings.Count(s, "-") >= 3 {
		return true
	}
	// служебные шифры ГРАНД-Сметы: 49-1, 49-2 и т.п.
	if reSimpleCode.MatchString(s) {
		return true
	}
	return false
}

// Parse читает xlsx-выгрузку ЛСР.
func Parse(path string) (*Estimate, error) {
	book, err := xlsx.Open(path)
	if err != nil {
		return nil, err
	}
	sh := book.Sheets[0]
	est := &Estimate{}

	row := 1
	for row <= 40 && row <= sh.MaxRow {
		t := sh.Text(1, row)
		switch {
		case strings.HasPrefix(t, "ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ"):
			est.LocNum = strings.TrimPrefix(t, "ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА) №")
			est.LocNum = strings.TrimSpace(est.LocNum)
		case t == "Сметная стоимость":
			est.Description = sh.Text(21, row)
		case strings.HasPrefix(t, "Наименование субъекта"):
			est.RegionName = sh.Text(7, row)
			est.RegionCode = regionCode(est.RegionName)
		case strings.HasPrefix(t, "Основание"):
			est.Basis = sh.Text(2, row)
		case t == "Составлен(а) в текущем уровне цен":
			est.CurrPriceDate = sh.Text(3, row)
		case strings.HasPrefix(t, "Наименование редакции сметных нормативов"):
			est.Approved = sh.Text(7, row)
		case strings.HasPrefix(t, "(наименование стройки)") && row > 1:
			est.Constr = sh.Text(1, row-1)
		case strings.HasPrefix(t, "(наименование объекта"):
			est.Object = sh.Text(1, row-1)
		}
		if t == "Составил:" {
			// ФИО не переносим (правило проекта), оставляем заглушку
		}
		row++
	}

	// Основная таблица: шапка в строках 35–38, данные с 39.
	var cur *Chapter
	var pos *Position
	var curRes *Resource

	for r := 39; r <= sh.MaxRow; r++ {
		a := sh.Text(1, r)
		b := sh.Text(2, r)
		c := sh.Text(3, r)
		if a == "" && b == "" && c == "" {
			continue
		}

		// Раздел
		if m := reChapter.FindStringSubmatch(a); m != nil {
			cur = &Chapter{Caption: m[2], Headers: nil}
			est.Chapters = append(est.Chapters, cur)
			pos = nil
			continue
		}

		// Итоги — не переносим
		if strings.HasPrefix(c, "Итого") || c == "ФОТ" || c == "Всего по позиции" ||
			strings.HasPrefix(c, "Итоги по") || a == "ВСЕГО по смете" {
			pos = nil
			continue
		}

		// Позиция: A число, B код
		if rePosition.MatchString(a) && isCode(b) {
			n, _ := strconv.Atoi(a)
			p := &Position{
				Number:  n,
				Code:    b,
				Caption: c,
				Units:   sh.Text(8, r),
				Row:     r,
			}
			p.Quantity, _ = sh.Num(9, r)
			if k, ok := sh.Num(10, r); ok && k != 1 {
				p.K = k
				p.HasK = true
			} else if ok {
				p.K = k
				p.HasK = true
			}
			if cur == nil {
				cur = &Chapter{Caption: "Раздел 1"}
				est.Chapters = append(est.Chapters, cur)
			}
			cur.Positions = append(cur.Positions, p)
			pos = p
			curRes = nil
			continue
		}

		// Заголовок подраздела: только A, без B и C
		if a != "" && b == "" && c == "" && pos == nil && cur != nil {
			cur.Headers = append(cur.Headers, a)
			continue
		}

		// Строка коэффициента: шифр в B (приказ) или в начале C (приложение),
		// значения вида «ОЗП=1,2; ТЗ=1,2» в тексте.
		if pos != nil && a == "" && c != "" && sh.Text(8, r) == "" && isKoeffRow(b, c) {
			if kf, ok := parseKoeff(b, c); ok {
				pos.Koeffs = append(pos.Koeffs, kf)
			}
			continue
		}

		// Ресурс: B — код ресурса, C — наименование, H — ед. (нитка «ОТ(ЗТ)» — агрегат, не ресурс)
		if pos != nil && b != "" && c != "" && sh.Text(8, r) != "" && sh.Text(8, r) != "%" && c != "ОТ(ЗТ)" {
			res := Resource{
				Kind:    resourceKind(b),
				Code:    b,
				Caption: c,
				Units:   sh.Text(8, r),
				Row:     r,
			}
			res.Quantity, _ = sh.Num(9, r)
			if k, ok := sh.Num(10, r); ok {
				res.K = k
				res.HasK = true
			}
			if v, ok := sh.Num(12, r); ok {
				res.Base = v
				res.HasBase = true
			}
			if v, ok := sh.Num(13, r); ok {
				res.Index = v
				res.HasIndex = true
			}
			if v, ok := sh.Num(14, r); ok {
				res.Curr = v
				res.HasCurr = true
			}
			pos.Res = append(pos.Res, res)
			curRes = &pos.Res[len(pos.Res)-1]
			continue
		}

		// Уточнение ресурса: B код, C текст, но единиц нет — уточняет предыдущий ресурс
		if pos != nil && curRes != nil && b != "" && c != "" && sh.Text(8, r) == "" {
			_ = r
			continue
		}
	}

	if est.LocNum == "" {
		return nil, fmt.Errorf("%s: не найден номер локального сметного расчёта", path)
	}
	if len(est.Chapters) == 0 {
		return nil, fmt.Errorf("%s: не найдено ни одного раздела", path)
	}
	total := 0
	for _, ch := range est.Chapters {
		total += len(ch.Positions)
		if len(ch.Positions) == 0 {
			est.Warnings = append(est.Warnings, "раздел без позиций: "+ch.Caption)
		}
	}
	if total == 0 {
		return nil, fmt.Errorf("%s: не найдено ни одной позиции", path)
	}
	return est, nil
}

// resourceKind определяет тип ресурса по шифру.
// Tzr     — труд рабочих (1-100-20);
// TzrMch  — труд машинистов (4-100-060), вкладывается в машину как Tzm2022;
// Mch     — машины и механизмы (91.xx.xx-xxx);
// Mat     — материалы (02.x / 01.x / ФСБЦ / ТЦ_).
func resourceKind(code string) string {
	switch {
	case strings.HasPrefix(code, "91."):
		return "Mch"
	case strings.HasPrefix(code, "4-"):
		return "TzrMch"
	case strings.HasPrefix(code, "1-"):
		return "Tzr"
	}
	return "Mat"
}

var reRegion = regexp.MustCompile(`(\d{2,3})\.`)

func regionCode(name string) string {
	if m := reRegion.FindStringSubmatch(name); m != nil {
		return m[1]
	}
	return ""
}

// isKoeffRow сообщает, является ли строка коэффициентом позиции.
func isKoeffRow(code, text string) bool {
	text = strings.TrimSpace(text)
	if text == "" || reKoeffPart.FindStringIndex(text) == nil {
		return false
	}
	if strings.HasPrefix(text, "Итого") || strings.HasPrefix(text, "Всего") ||
		strings.HasPrefix(text, "ФОТ") {
		return false
	}
	if code != "" {
		return reKoeffCode.MatchString(code)
	}
	return reKoeffHead.MatchString(text)
}

// reKoeffCode — шифр коэффициента в колонке B.
var reKoeffCode = regexp.MustCompile(`^(Приказ|Прил\.|421/пр|Пр/|МДС|п\.\d)`)

var (
	reKoeffHead = regexp.MustCompile(`^(Прил\.|421/пр_|Приказ от|Пр/|МДС|п\.\d)`)
	reKoeffPart = regexp.MustCompile(`([А-ЯЁA-Z]{2,4})\s*=\s*(\d+(?:[,.]\d+)?)`)
)

// parseKoeff разбирает строку «<текст> ОЗП=1,2; ЭМ=1,2 к расх.; ЗПМ=1,2; ТЗ=1,2; ТЗМ=1,2».
func parseKoeff(code, s string) (Koeff, bool) {
	kp := Koeff{Raw: strings.TrimSpace(s)}
	m := reKoeffPart.FindStringIndex(s)
	if m == nil {
		return kp, false
	}
	kp.Caption = strings.TrimSpace(s[:m[0]])
	values := s[m[0]:]

	kp.Code = strings.TrimSpace(code)
	if kp.Code == "" {
		if i := strings.Index(kp.Caption, "_"); i > 0 {
			kp.Code = strings.TrimSpace(kp.Caption[:i])
		} else if i := strings.Index(kp.Caption, " "); i > 0 {
			kp.Code = strings.TrimSpace(kp.Caption[:i])
		}
	}

	num := func(raw string) float64 {
		f, err := strconv.ParseFloat(strings.ReplaceAll(raw, ",", "."), 64)
		if err != nil {
			return 0
		}
		return f
	}

	for _, part := range strings.Split(values, ";") {
		part = strings.TrimSpace(part)
		g := reKoeffPart.FindStringSubmatch(part)
		if g == nil {
			continue
		}
		v := num(g[2])
		switch g[1] {
		case "ОЗП", "ОЗ":
			kp.OZP, kp.HasOZP = v, true
		case "ЭМ":
			kp.EM, kp.HasEM = v, true
			kp.EMRash = strings.Contains(part, "расх")
		case "ЗПМ", "ЗПМЗ":
			kp.ZPM, kp.HasZPM = v, true
		case "ТЗ", "ОТ":
			kp.TZ, kp.HasTZ = v, true
		case "ТЗМ":
			kp.TZM, kp.HasTZM = v, true
		}
	}
	return kp, kp.HasOZP || kp.HasEM || kp.HasZPM || kp.HasTZ || kp.HasTZM
}
