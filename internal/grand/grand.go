// Package grand — генератор XML ГРАНД-Сметы (формат внутреннего обмена .xml).
package grand

import (
	"bytes"
	"fmt"
	"regexp"
	"sort"
	"strings"

	"xlsx2grand/internal/smeta"
	"xlsx2grand/internal/vor"
)

// Join — результат сопоставления ВОР и сметы.
type JoinItem struct {
	VOR      *vor.Item
	Position *smeta.Position // nil, если сопоставить не удалось
	Note     string
}

// Options — параметры генерации.
type Options struct {
	Generator      string // "GrandSmeta"
	ProgramVersion string // "15.3.3.54868"
	DocTypeGUID    string // "{2B0470FD-477C-4359-9F34-EEBE36B7D340}"
	Composer       string // "—"
	CurrPriceDate  string // ДД.ММ.ГГГГ
}

// Build собирает XML.
func Build(est *smeta.Estimate, v *vor.VOR, opts Options) ([]byte, []*JoinItem) {
	if opts.Generator == "" {
		opts.Generator = "GrandSmeta"
	}
	if opts.ProgramVersion == "" {
		opts.ProgramVersion = "15.3.3.54868"
	}
	if opts.DocTypeGUID == "" {
		opts.DocTypeGUID = "{2B0470FD-477C-4359-9F34-EEBE36B7D340}"
	}
	if opts.Composer == "" {
		opts.Composer = "—"
	}

	positions := map[string]*smeta.Position{}
	for _, ch := range est.Chapters {
		for _, p := range ch.Positions {
			positions[fmt.Sprint(p.Number)] = p
		}
	}

	// Общие коэффициенты (охранная зона, стеснённые условия) ГРАНД-Смета
	// держит одним блоком на документе, а не копией в каждой позиции.
	global, perPos := splitKoeffs(est)

	b := &buf{}
	b.line(`<?xml version="1.0" encoding="windows-1251"?>`)
	b.rawf(`<Document ProgramVersion=%s Generator=%s DocumentType=%s MaxSysID="%d">`,
		attr(opts.ProgramVersion), attr(opts.Generator), attr(opts.DocTypeGUID), len(vorItems(v))+4)

	b.rawf(`  <Properties LocNum=%s Object=%s Constr=%s Description=%s/>`,
		attr(est.LocNum), attr(est.Object), attr(est.Constr), attr(est.Description))

	// Подписи
	b.line(`  <GsDocSignatures>`)
	b.rawf(`    <Item Caption="Составил" ID="300" Value=%s/>`, attr(opts.Composer))
	b.rawf(`    <Item Caption="Основание" ID="100" Value=%s/>`, attr(est.Basis))
	b.line(`  </GsDocSignatures>`)

	// Даты
	if d := parseDate(opts.CurrPriceDate); d != "" {
		b.rawf(`  <DocDates СreationDate=%s ApprovalDate=%s BasePriceDate="01.01.2001" CurrPriceDate=%s/>`,
			attr(d), attr(d), attr(d))
	} else {
		b.line(`  <DocDates BasePriceDate="01.01.2001"/>`)
	}

	b.line(`  <RegionalK Options="Percent Base"/>`)
	b.line(`  <TerZoneK Caption="Территориальная поправка к базе 2001г" Options="Base PzAll EmAll"/>`)
	writeGlobalKoeffs(b, global)

	if code := est.RegionCode; code != "" {
		b.rawf(`  <RegionInfo RegionName=%s AdmRegionCode=%s AdmRegionZone=%s/>`,
			attr(est.Approved), attr(code), attr(est.RegionName))
	}

	b.line(`  <Parameters Options="SeparateFot 2022Mode" Mode2020Order="2024" BasePrices="2001" BaseCalcVrs="Vr2001" TzDigits="2" MatDigits="4"/>`)

	// Разделы
	b.line(`  <Chapters>`)
	total := 0
	sysID := 1
	var joins []*JoinItem

	for ci, ch := range v.Sections {
		b.rawf(`    <Chapter Caption=%s SysID="%d" IndexCode="1">`, attr(ch.Caption), ci+1)

		// Подзаголовок выводится один раз — там, где меняется у строк.
		curHeader := ""

		for _, it := range ch.Items {
			if it.Header != "" && it.Header != curHeader {
				b.rawf(`      <Header Caption=%s/>`, attr(it.Header))
				curHeader = it.Header
			}
			pos := positions[it.NumLSR]
			ji := &JoinItem{VOR: it, Position: pos}
			joins = append(joins, ji)
			if pos == nil {
				ji.Note = "в смете нет позиции с таким номером"
			}

			num := it.Num
			if num == 0 {
				num = total + 1
			}
			caption := it.Name
			code := ""
			units := it.Unit
			qty := fmtNum(it.Quantity)
			fx := it.Formula
			vrb := ""
			dbComment := ""
			indexCode := ""

			if pos != nil {
				if caption == "" {
					caption = pos.Caption
				}
				code = pos.Code
				units = pos.Units
				_ = pos.Quantity
				dbComment = est.Approved
			}

			attrs := []string{
				`Caption=` + attr(caption),
				fmt.Sprintf(`SysID="%d"`, sysID),
				fmt.Sprintf(`Number="%d"`, num),
			}
			if code != "" {
				attrs = append(attrs, `Code=`+attr(code))
			}
			attrs = append(attrs, `Units=`+attr(units), `PriceLevel="2001"`)
			if dbComment != "" {
				attrs = append(attrs, `DBComment=`+attr(dbComment))
			}
			attrs = append(attrs, `Quantity=`+attr(qty), `PzSync="Yes"`)
			if vrb != "" {
				attrs = append(attrs, `Vr2001=`+attr(vrb))
			}
			_ = indexCode

			// Служебные шифры (49-1) в ресурсы не выводим.
			if pos != nil && pos.Code != "" {
				attrs = append(attrs, `Vr2001=""`)
				attrs[len(attrs)-1] = ""
			}
			b.rawf(`      <Position %s>`, strings.Join(nonEmpty(attrs), " "))
			b.rawf(`        <Quantity Fx=%s KUnit="1" Result=%s/>`, attr(fx), attr(fmtNum(it.Quantity)))

			if pos != nil {
				writeResources(b, pos.Res)
				writeKoeffs(b, perPos[pos.Number])
			}
			b.line(`      </Position>`)
			sysID++
			total++
		}

		b.line(`    </Chapter>`)
	}
	b.line(`  </Chapters>`)
	b.line(`  <ReportOptions Options="Grouping" Kk="1">`)
	b.line(`    <RangingRates Mch="20" Mat="20"/>`)
	b.line(`  </ReportOptions>`)
	b.line(`</Document>`)
	return encodeCP1251(b.String()), joins
}

func vorItems(v *vor.VOR) []*vor.Item {
	var out []*vor.Item
	for _, s := range v.Sections {
		out = append(out, s.Items...)
	}
	return out
}

func orDefault(s, def string) string {
	if s == "" {
		return def
	}
	return s
}

var xmlEsc = strings.NewReplacer(
	"&", "&amp;",
	"<", "&lt;",
	">", "&gt;",
	`"`, "&quot;",
	"\n", "&#10;",
	"\r", "&#13;",
	"\t", "&#9;",
)

func esc(s string) string {
	s, _ = sanitize(s)
	return xmlEsc.Replace(s)
}

func attr(s string) string { return `"` + esc(s) + `"` }

// fmtNum печатает число с запятой, без экспоненты и лишних нулей.
func fmtNum(f float64) string {
	s := formatFloat(f)
	s = strings.ReplaceAll(s, ".", ",")
	return s
}

var reChapterNum = regexp.MustCompile(`(\d{2})\.(\d{2})\.(\d{4})`)

func parseDate(s string) string {
	m := reChapterNum.FindStringSubmatch(strings.TrimSpace(s))
	if m == nil {
		return ""
	}
	return m[0]
}

// Дополнительные функции сортировки нужны для стабильного вывода.
func sortedKeys(m map[string]*smeta.Position) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	return keys
}

type buf struct {
	bytes.Buffer
}

func (b *buf) line(s string) { b.WriteString(s); b.WriteString("\r\n") }
func (b *buf) rawf(format string, args ...any) {
	fmt.Fprintf(&b.Buffer, format, args...)
	b.WriteString("\r\n")
}
func (b *buf) bytes() []byte { return b.Buffer.Bytes() }

// writeResources выводит ресурсы позиции в порядке, требуемом схемой
// OpenXML_LS: Tzr -> Tzm -> Mch -> Mat. Труд машинистов (4-100-xxx)
// не идёт отдельной строкой, а вкладывается в машину как <Tzm2022>.
func writeResources(b *buf, res []smeta.Resource) {
	if len(res) == 0 {
		return
	}
	// Индекс труда машинистов по количеству, чтобы связать с машиной.
	mchMech := map[float64]int{}
	for i := range res {
		if res[i].Kind == "TzrMch" {
			if _, seen := mchMech[res[i].Quantity]; !seen {
				mchMech[res[i].Quantity] = i
			}
		}
	}

	b.line(`        <Resources>`)
	for _, r := range res {
		if r.Kind != "Tzr" {
			continue
		}
		b.rawf(`          <Tzr Caption=%s Code=%s Units=%s Quantity=%s>`,
			attr(r.Caption), attr(r.Code), attr(orDefault(r.Units, "чел.-ч")), attr(fmtNum(r.Quantity)))
		if r.HasCurr {
			b.rawf(`            <PriceCurr Value=%s/>`, attr(fmtNum(r.Curr)))
		}
		b.line(`          </Tzr>`)
	}

	// Tzm — суммарные затраты труда машинистов по позиции (код "2").
	var tzm float64
	hasTzm := false
	for _, r := range res {
		if r.Kind == "TzrMch" {
			tzm += r.Quantity
			hasTzm = true
		}
	}
	if hasTzm {
		b.rawf(`          <Tzm Caption="Затраты труда машинистов" Code="2" Units="чел.-ч" Quantity=%s/>`,
			attr(fmtNum(tzm)))
	}

	usedTzrMch := map[int]bool{}
	mchIdx := 0
	for _, r := range res {
		if r.Kind != "Mch" {
			continue
		}
		mchIdx++
		b.rawf(`          <Mch Caption=%s Code=%s Units=%s Quantity=%s>`,
			attr(r.Caption), attr(r.Code), attr(orDefault(r.Units, "маш.-ч")), attr(fmtNum(r.Quantity)))
		if r.HasBase {
			b.rawf(`            <PriceBase Value=%s/>`, attr(fmtNum(r.Base)))
		}
		if r.HasCurr {
			b.rawf(`            <PriceCurr Value=%s/>`, attr(fmtNum(r.Curr)))
		}
		ti, ok := mchMech[r.Quantity]
		if !ok {
			ti, ok = takeNearestTzrMch(res, usedTzrMch, mchIdx)
		}
		if ok && !usedTzrMch[ti] {
			usedTzrMch[ti] = true
			tm := res[ti]
			class := machineClass(tm.Caption)
			b.rawf(`            <Tzm2022 WorkClass=%s Quantity="1" Price=%s/>`,
				attr(class), attr(fmtNum(tm.Curr)))
		}
		b.line(`          </Mch>`)
	}

	for _, r := range res {
		if r.Kind != "Mat" {
			continue
		}
		b.rawf(`          <Mat Caption=%s Code=%s Units=%s Quantity=%s>`,
			attr(r.Caption), attr(r.Code), attr(r.Units), attr(fmtNum(r.Quantity)))
		if r.HasBase {
			b.rawf(`            <PriceBase Value=%s/>`, attr(fmtNum(r.Base)))
		}
		if r.HasCurr {
			b.rawf(`            <PriceCurr Value=%s/>`, attr(fmtNum(r.Curr)))
		}
		b.line(`          </Mat>`)
	}
	b.line(`        </Resources>`)
}

var reWorkClass = regexp.MustCompile(`(\d+(?:[,.]\d+)?)\s*$`)

// machineClass достаёт разряд машинистов из «Средний разряд машинистов 6».
func machineClass(caption string) string {
	m := reWorkClass.FindStringSubmatch(strings.TrimSpace(caption))
	if m == nil {
		return ""
	}
	return strings.ReplaceAll(m[1], ".", ",")
}

// takeNearestTzrMch находит незанятый труд машинистов, ближайший к машине
// по порядку следования (когда количества не совпали).
func takeNearestTzrMch(res []smeta.Resource, used map[int]bool, mchIdx int) (int, bool) {
	seen := 0
	for i := range res {
		if res[i].Kind != "Mch" {
			continue
		}
		seen++
		if seen != mchIdx {
			continue
		}
		// ищем первый незанятый TzrMch после этой машины
		for j := i + 1; j < len(res); j++ {
			if res[j].Kind == "TzrMch" && !used[j] {
				return j, true
			}
			if res[j].Kind == "Mch" {
				break
			}
		}
		break
	}
	return 0, false
}

// writeKoeffs выводит коэффициенты позиции.
func writeKoeffs(b *buf, kfs []smeta.Koeff) {
	if len(kfs) == 0 {
		return
	}
	b.line(`        <Koefficients>`)
	for i, k := range kfs {
		opts := "Base Curr EmAll EmQty OzpTz ZpmTz"
		if k.EMRash {
			opts += " AddOn10"
		}
		attrs := []string{}
		if k.Caption != "" {
			attrs = append(attrs, `Caption=`+attr(k.Caption))
		}
		if k.Code != "" {
			attrs = append(attrs, `Code=`+attr(k.Code))
		}
		attrs = append(attrs, `Options=`+attr(opts))
		if k.HasOZP {
			attrs = append(attrs, `Value_OZ=`+attr(fmtNum(k.OZP)))
		}
		if k.HasEM {
			attrs = append(attrs, `Value_EM=`+attr(fmtNum(k.EM)))
		}
		if k.HasZPM {
			attrs = append(attrs, `Value_MT=`+attr(fmtNum(k.ZPM)))
		}
		attrs = append(attrs, fmt.Sprintf(`Level="%d"`, i+1))
		b.rawf(`          <K %s/>`, strings.Join(attrs, " "))
	}
	b.line(`        </Koefficients>`)
}

// nonEmpty отбрасывает пустые фрагменты атрибутов.
func nonEmpty(attrs []string) []string {
	out := attrs[:0]
	for _, a := range attrs {
		if strings.TrimSpace(a) != "" {
			out = append(out, a)
		}
	}
	return out
}

// splitKoeffs делит коэффициенты на общие (код повторяется у большинства
// позиций) и позиционные. Общие ГРАНД-Смета хранит одним блоком документа.
func splitKoeffs(est *smeta.Estimate) ([]smeta.Koeff, map[int][]smeta.Koeff) {
	type rec struct {
		kf    smeta.Koeff
		count int
		order int
	}
	seen := map[string]*rec{}
	order := 0
	total := 0
	for _, ch := range est.Chapters {
		for _, p := range ch.Positions {
			total++
			for _, k := range p.Koeffs {
				key := k.Code
				if key == "" {
					key = k.Caption
				}
				r, ok := seen[key]
				if !ok {
					order++
					r = &rec{kf: k, order: order}
					seen[key] = r
				}
				r.count++
			}
		}
	}

	global := make([]smeta.Koeff, 0, len(seen))
	isGlobal := map[string]bool{}
	for key, r := range seen {
		// Коэффициент общий, если применён не менее чем к трети позиций.
		if total > 0 && r.count*3 >= total {
			global = append(global, r.kf)
			isGlobal[key] = true
		}
	}
	sort.SliceStable(global, func(i, j int) bool {
		return seen[keyOf(global[i])].order < seen[keyOf(global[j])].order
	})

	perPos := map[int][]smeta.Koeff{}
	for _, ch := range est.Chapters {
		for _, p := range ch.Positions {
			for _, k := range p.Koeffs {
				if isGlobal[keyOf(k)] {
					continue
				}
				perPos[p.Number] = append(perPos[p.Number], k)
			}
		}
	}
	return global, perPos
}

func keyOf(k smeta.Koeff) string {
	if k.Code != "" {
		return k.Code
	}
	return k.Caption
}

// writeGlobalKoeffs выводит документный блок общих коэффициентов.
func writeGlobalKoeffs(b *buf, kfs []smeta.Koeff) {
	if len(kfs) == 0 {
		b.line(`  <Koefficients/>`)
		return
	}
	b.line(`  <Koefficients>`)
	for i, k := range kfs {
		opts := "InPos Base Curr EmAll MatQty EmQty OzpTz ZpmTz AddOn10"
		attrs := []string{}
		if k.Caption != "" {
			attrs = append(attrs, `Caption=`+attr(k.Caption))
		}
		if k.Code != "" {
			attrs = append(attrs, `Code=`+attr(k.Code))
		}
		attrs = append(attrs, `Options=`+attr(opts))
		if k.HasOZP {
			attrs = append(attrs, `Value_OZ=`+attr(fmtNum(k.OZP)))
		}
		if k.HasEM {
			attrs = append(attrs, `Value_EM=`+attr(fmtNum(k.EM)))
		}
		attrs = append(attrs, fmt.Sprintf(`Level="%d"`, i+1))
		b.rawf(`    <K %s/>`, strings.Join(attrs, " "))
	}
	b.line(`  </Koefficients>`)
}
