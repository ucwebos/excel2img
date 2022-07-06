package lib

import (
	"github.com/golang/freetype/truetype"
	"image/color"
	"unicode"
)

const defaultSize = 12

type ICell struct {
	Axis    string
	Hide    bool
	MrMain  bool
	MrTypeX bool
	Value   string
	Style   *Style
	Width   int
	Height  int
}

func (c *ICell) getBgColor() color.Color {
	cr := c.Style.Fill.FgColor
	if cr == "" {
		return color.White
	}
	return colorFromStr(cr)
}

func (c *ICell) getFontTT() *truetype.Font {
	family := c.Style.Font.Name
	if c.Style.Font.Bold {
		family = family + "_bold"
		if ft, ok := fontTTs[family]; ok {
			return ft
		}
		return fontTTs["微软雅黑_bold"]
	}
	if ft, ok := fontTTs[family]; ok {
		return ft
	}
	return fontTTs["微软雅黑"]
}

func (c *ICell) getFontColor() color.Color {
	cr := c.Style.Font.Color
	if cr == "" {
		return color.Black
	}
	return colorFromStr(cr)
}

func (c *ICell) getBorderLeftColor() color.Color {
	cr := c.Style.Border.Left
	if cr == "" {
		return color.Black
	}
	return colorFromStr(cr)
}
func (c *ICell) getBorderRightColor() color.Color {
	cr := c.Style.Border.Right
	if cr == "" {
		return color.Black
	}
	return colorFromStr(cr)
}

func (c *ICell) getBorderTopColor() color.Color {
	cr := c.Style.Border.Top
	if cr == "" {
		return color.Black
	}
	return colorFromStr(cr)
}

func (c *ICell) getBorderBottomColor() color.Color {
	cr := c.Style.Border.Bottom
	if cr == "" {
		return color.Black
	}
	return colorFromStr(cr)
}

func (c *ICell) getSize() int {
	size := c.Style.Font.Size
	if size == 0 {
		return defaultSize
	}
	return int(size)
}

func (c *ICell) getBeginPY() (y int) {
	y = 0
	switch c.Style.Alignment.Vertical {
	case "center":
		y = c.Height/2 + c.getSize()
	case "top":
		y = int(float64(c.getSize()) * 2)
	case "bottom", "":
		y = c.Height - c.getSize()
	}
	return y
}

func (c *ICell) getValWidth() (w int) {
	w = 0
	rs := []rune(c.Value)
	for _, s := range rs {
		if unicode.Is(unicode.Han, s) {
			w += int(float64(c.getSize()) * 3)
			continue
		}
		if unicode.IsPunct(s) {
			w += int(float64(c.getSize()) * 0.5)
			continue
		}
		w += c.getSize() * 2
	}
	return
}

func (c *ICell) getWh() (w, h int) {
	w, h = 20, 20
	w += c.getValWidth()
	h += int(2 * float64(c.getSize()))
	return
}
