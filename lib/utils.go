package lib

import (
	"fmt"
	"image/color"
	"strconv"
	"strings"
)

func colorFromStr(colorStr string) color.Color {
	// ARGB 转 RGBA
	if len(colorStr) == 8 {
		colorStr = colorStr[2:] + colorStr[0:2]
	}
	r, g, b, a := parseHexColor(colorStr)
	cr := color.RGBA{
		R: uint8(r),
		G: uint8(g),
		B: uint8(b),
		A: uint8(a),
	}
	return cr
}

func parseHexColor(x string) (r, g, b, a int) {
	x = strings.TrimPrefix(x, "#")
	a = 255
	if len(x) == 3 {
		format := "%1x%1x%1x"
		_, _ = fmt.Sscanf(x, format, &r, &g, &b)
		r |= r << 4
		g |= g << 4
		b |= b << 4
	}
	if len(x) == 6 {
		format := "%02x%02x%02x"
		_, _ = fmt.Sscanf(x, format, &r, &g, &b)
	}
	if len(x) == 8 {
		format := "%02x%02x%02x%02x"
		_, _ = fmt.Sscanf(x, format, &r, &g, &b, &a)
	}
	return
}

func IsNum(s string) bool {
	_, err := strconv.ParseFloat(s, 64)
	return err == nil
}
