// Функции форматирования чисел для XML ГРАНД-Сметы.
package grand

import (
	"math"
	"strconv"
)

// formatFloat печатает число в десятичном виде без экспоненты.
func formatFloat(f float64) string {
	if math.IsNaN(f) || math.IsInf(f, 0) {
		return "0"
	}
	if f == math.Trunc(f) && math.Abs(f) < 1e15 {
		return strconv.FormatInt(int64(f), 10)
	}
	return strconv.FormatFloat(f, 'f', -1, 64)
}
