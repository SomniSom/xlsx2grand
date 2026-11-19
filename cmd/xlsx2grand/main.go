// Команда xlsx2grand конвертирует выгрузки xlsx (ВОР и локальная смета)
// в XML ГРАНД-Сметы.
package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"xlsx2grand/internal/grand"
	"xlsx2grand/internal/smeta"
	"xlsx2grand/internal/vor"
)

func main() {
	var (
		vorPath   = flag.String("vor", "", "xlsx ведомости объёмов работ")
		smetaPath = flag.String("smeta", "", "xlsx локального сметного расчёта")
		outPath   = flag.String("out", "", "выходной XML (по умолчанию рядом с ВОР)")
		version   = flag.String("version", "15.3.3.54868", "ProgramVersion для ГРАНД-Сметы")
		composer  = flag.String("composer", "—", "подпись составителя")
	)
	flag.Parse()

	if *vorPath == "" || *smetaPath == "" {
		fmt.Fprintln(os.Stderr, "нужно указать -vor и -smeta")
		flag.Usage()
		os.Exit(2)
	}
	if *outPath == "" {
		base := strings.TrimSuffix(filepath.Base(*vorPath), filepath.Ext(*vorPath))
		*outPath = filepath.Join(filepath.Dir(*vorPath), base+".xml")
	}

	v, err := vor.Parse(*vorPath)
	if err != nil {
		fatal("ВОР", err)
	}
	est, err := smeta.Parse(*smetaPath)
	if err != nil {
		fatal("смета", err)
	}

	opts := grand.Options{
		ProgramVersion: *version,
		Composer:       *composer,
		CurrPriceDate:  est.CurrPriceDate,
	}
	data, joins := grand.Build(est, v, opts)

	if err := os.WriteFile(*outPath, data, 0o644); err != nil {
		fatal("запись XML", err)
	}

	report(v, est, joins)
	fmt.Printf("XML записан: %s\n", *outPath)
}

func report(v *vor.VOR, est *smeta.Estimate, joins []*grand.JoinItem) {
	fmt.Printf("ВОР: %d разделов, %d строк\n", len(v.Sections), len(joins))
	fmt.Printf("Смета: ЛСР %s, %d разделов\n", est.LocNum, len(est.Chapters))
	missing := 0
	for _, j := range joins {
		if j.Note != "" {
			missing++
			fmt.Printf("  ! строка ВОР %d (%s): %s\n", j.VOR.Num, truncate(j.VOR.Name, 50), j.Note)
		}
	}
	fmt.Printf("Не сопоставлено со сметой: %d\n", missing)

	used := map[int]bool{}
	for _, j := range joins {
		if j.Position != nil {
			used[j.Position.Number] = true
		}
	}
	orphans := 0
	for _, ch := range est.Chapters {
		for _, p := range ch.Positions {
			if !used[p.Number] {
				orphans++
			}
		}
	}
	fmt.Printf("Позиций сметы без строки в ВОР: %d\n", orphans)

	sumVOR, sumSmeta := 0.0, 0.0
	for _, j := range joins {
		sumVOR += j.VOR.Quantity
	}
	for _, ch := range est.Chapters {
		for _, p := range ch.Positions {
			sumSmeta += p.Quantity
		}
	}
	fmt.Printf("Сумма количества ВОР: %.6f, сумма количества сметы: %.6f\n", sumVOR, sumSmeta)
}

func truncate(s string, n int) string {
	r := []rune(s)
	if len(r) <= n {
		return s
	}
	return string(r[:n]) + "…"
}

func fatal(what string, err error) {
	fmt.Fprintf(os.Stderr, "ошибка (%s): %v\n", what, err)
	os.Exit(1)
}
