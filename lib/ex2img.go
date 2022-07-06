package lib

import (
	"bufio"
	"embed"
	"fmt"
	"github.com/golang/freetype"
	"github.com/golang/freetype/truetype"
	"github.com/xuri/excelize/v2"
	"golang.org/x/image/draw"
	"golang.org/x/image/font"
	"image"
	"image/color"
	"image/png"
	"os"
	"strings"
)

const colAxis = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
const colAxisLen = 26

var fontTTs = map[string]*truetype.Font{}

func Init(fonts embed.FS) error {
	var fontFiles = map[string]string{
		"微软雅黑":      "msyh.ttf",
		"微软雅黑_bold": "msyh_bold.ttf",
		"宋体":        "simsun.ttf",
		"宋体_bold":   "simsun_bold.ttf",
		"黑体":        "simhei.ttf",
	}
	for s, fontFile := range fontFiles {
		buf, err := fonts.ReadFile("fonts/" + fontFile)
		if err != nil {
			return err
		}
		ft, err := freetype.ParseFont(buf)
		if err != nil {
			return err
		}
		fontTTs[s] = ft
	}
	return nil
}

type Ex2Img struct {
	dWidth  int
	dHeight int
	mergeMG *MergeMG
}

// DrawExcelToPngFile  转换excel存储PNG图片到磁盘
func (d *Ex2Img) DrawExcelToPngFile(file *excelize.File, outPngName string) error {
	rgba, err := d.DrawExcel(file)
	if err != nil {
		return err
	}
	return d.save(outPngName, rgba)
}

// DrawExcel 转换excel返回image.RGBA 可以按需要转成各种图片
func (d *Ex2Img) DrawExcel(file *excelize.File) (rgba *image.RGBA, err error) {
	// 获取合并单元格
	sheet := file.GetSheetName(0)
	mergeCells, err := file.GetMergeCells(sheet)
	if err != nil {
		return
	}
	d.mergeMG = NewMergeMG(mergeCells)
	// 解析数据
	rows, xLen, err := d.parseRows(file)

	var (
		wMap = make(map[int]int)
		hMap = make(map[int]int)
	)
	// parse item width height
	for i := 0; i < xLen; i++ {
		if _, ok := wMap[i]; !ok {
			wMap[i] = 0
		}
		for j, row := range rows {
			if _, ok := hMap[j]; !ok {
				hMap[j] = 0
			}
			if len(row) > i {
				cell := row[i]
				x, y := cell.getWh()
				if cell.MrMain {
					if cell.MrTypeX {
						if y > hMap[j] {
							hMap[j] = y
						}
					} else {
						if x > wMap[i] {
							wMap[i] = x
						}
					}
					continue
				}
				if x > wMap[i] {
					wMap[i] = x
				}
				if y > hMap[j] {
					hMap[j] = y
				}
			}
		}
	}
	for i, row := range rows {
		for j := range row {
			rows[i][j].Width = wMap[j]
			rows[i][j].Height = hMap[i]
		}
	}
	// set dWidth, dHeight
	for _, w := range wMap {
		d.dWidth += w
	}
	for _, h := range hMap {
		d.dHeight += h
	}
	return d.draw(rows), nil
}

func (d *Ex2Img) GetStyle(file *excelize.File, styleID int) *Style {
	cs := file.Styles.CellXfs.Xf[styleID]
	themes := file.Theme.ThemeElements
	var (
		ft  = file.Styles.Fonts.Font[*cs.FontID]
		fi  = file.Styles.Fills.Fill[*cs.FillID]
		br  = file.Styles.Borders.Border[*cs.BorderID]
		agt = cs.Alignment
	)

	style := &Style{
		Border:         Border{},
		Fill:           Fill{},
		Font:           Font{},
		ApplyBorder:    false,
		ApplyFill:      false,
		ApplyFont:      false,
		ApplyAlignment: false,
		Alignment:      Alignment{},
	}
	if br != nil {
		style.Border = Border{
			Left:   br.Left.Style,
			Right:  br.Right.Style,
			Top:    br.Top.Style,
			Bottom: br.Bottom.Style,
		}
	}
	if agt != nil {
		style.Alignment = Alignment{
			Horizontal:   agt.Horizontal,
			Indent:       agt.Indent,
			ShrinkToFit:  agt.ShrinkToFit,
			TextRotation: agt.TextRotation,
			Vertical:     agt.Vertical,
			WrapText:     agt.WrapText,
		}
	}

	if br.Left.Color != nil {
		style.Border.LeftColor = br.Left.Color.RGB
	}
	if br.Right.Color != nil {
		style.Border.RightColor = br.Right.Color.RGB
	}
	if br.Top.Color != nil {
		style.Border.TopColor = br.Top.Color.RGB
	}
	if br.Bottom.Color != nil {
		style.Border.BottomColor = br.Bottom.Color.RGB
	}
	if fi.PatternFill != nil {
		style.Fill.PatternType = fi.PatternFill.PatternType
		if fi.PatternFill.BgColor != nil {
			style.Fill.BgColor = fi.PatternFill.BgColor.RGB
		}

		if fi.PatternFill.FgColor != nil {
			style.Fill.FgColor = fi.PatternFill.FgColor.RGB
			if style.Fill.FgColor == "" && fi.PatternFill.FgColor.Theme != nil {
				themeID := *fi.PatternFill.FgColor.Theme
				children := themes.ClrScheme.Children
				if themeID < 4 {
					dklt := map[int]string{
						0: children[1].SysClr.LastClr,
						1: children[0].SysClr.LastClr,
						2: *children[3].SrgbClr.Val,
						3: *children[2].SrgbClr.Val,
					}
					style.Fill.FgColor = strings.TrimPrefix(excelize.ThemeColor(dklt[themeID], fi.PatternFill.FgColor.Tint), "FF")
				} else {
					srgbClr := children[themeID].SrgbClr.Val
					style.Fill.FgColor = strings.TrimPrefix(excelize.ThemeColor(*srgbClr, fi.PatternFill.FgColor.Tint), "FF")
				}

			}
		}
	}
	if ft.Sz != nil {
		v := ft.Sz.Val
		style.Font.Size = *v
	}
	if ft.Name != nil {
		v := ft.Name.Val
		style.Font.Name = *v
	}
	if ft.Color != nil {
		style.Font.Color = ft.Color.RGB
		if style.Font.Color == "" {
			if ft.Color.Theme != nil {
				themeID := *ft.Color.Theme
				if themeID < 4 {
					children := themes.ClrScheme.Children
					dklt := map[int]string{
						0: children[1].SysClr.LastClr,
						1: children[0].SysClr.LastClr,
						2: *children[3].SrgbClr.Val,
						3: *children[2].SrgbClr.Val,
					}
					style.Font.Color = strings.TrimPrefix(excelize.ThemeColor(dklt[themeID], ft.Color.Tint), "FF")
				}
			}
		}
	}
	if ft.B != nil {
		v := ft.B.Val
		style.Font.Bold = *v
	}
	if ft.I != nil {
		v := ft.I.Val
		style.Font.Italic = *v
	}
	if ft.U != nil {
		style.Font.Underline = true
	}
	if ft.Strike != nil {
		v := ft.Strike.Val
		style.Font.Strike = *v
	}
	return style
}

func (d *Ex2Img) FormatNum(file *excelize.File, styleID int, val string) (string, error) {
	cs := file.Styles.CellXfs.Xf[styleID]
	if cs.NumFmtID == nil {
		return val, nil
	}
	numFmtID := *cs.NumFmtID
	if file.Styles.NumFmts != nil {
		for _, numFmt := range file.Styles.NumFmts.NumFmt {
			if numFmt.NumFmtID == numFmtID {
				pNumFt := parseFullNumberFormatString(numFmt.FormatCode)
				return pNumFt.formatNumericCell(val)
			}
		}
	}
	return val, nil
}

func (d *Ex2Img) parseRows(file *excelize.File) (rows [][]*ICell, xLen int, err error) {
	sheet1 := file.GetSheetList()[0]
	rows = make([][]*ICell, 0)
	xLen = 0
	//opts := excelize.Options{
	//	RawCellValue:      true,
	//}
	data, err := file.GetRows(sheet1)
	if err != nil {
		return
	}
	for i, datum := range data {
		row := make([]*ICell, 0)
		iLen := 0
		for j := range datum {
			preFix := ""
			if j/colAxisLen >= 1 {
				preFix = string(colAxis[j/colAxisLen-1])
			}
			axis := fmt.Sprintf("%s%s%d", preFix, string(colAxis[j%colAxisLen]), i+1)
			val, _ := file.GetCellValue(sheet1, axis)
			styleID, err := file.GetCellStyle(sheet1, axis)
			if err != nil {
				fmt.Printf("file.GetCellStyle(%s, %s)  err %v\n", sheet1, axis, err)
				continue
			}
			if IsNum(val) {
				cType, err := file.GetCellType(sheet1, axis)
				if err == nil {
					if cType == excelize.CellTypeString && strings.HasSuffix(val, ".00") {
						val = strings.TrimSuffix(val, ".00")
					}
				}
				val, _ = d.FormatNum(file, styleID, val)
			}
			iCell := &ICell{
				Axis:  axis,
				Value: val,
				Style: d.GetStyle(file, styleID),
			}
			// 判断是被合并单元格
			if mgr, ok := d.mergeMG.HideCells[iCell.Axis]; ok {
				iCell.Hide = true
				iCell.Value = ""
				mgr.MrCells = append(mgr.MrCells, iCell)
			}
			// 判断是被合并单元格@Y
			if mgr, ok := d.mergeMG.HideCellsY[iCell.Axis]; ok {
				iCell.Hide = true
				iCell.Value = ""
				mgr.MrCellsY = append(mgr.MrCellsY, iCell)
			}
			if v, ok := d.mergeMG.MainCells[iCell.Axis]; ok {
				iCell.MrMain = true
				iCell.MrTypeX = v.TypeX
			}
			row = append(row, iCell)
			iLen += 1
		}
		rows = append(rows, row)
		if iLen > xLen {
			xLen = iLen
		}
	}
	// 补齐无数据单元格
	for i, row := range rows {
		if len(row) < xLen {
			for j := len(row); j < xLen; j++ {
				preFix := ""
				if j/colAxisLen >= 1 {
					preFix = string(colAxis[j/colAxisLen-1])
				}
				axis := fmt.Sprintf("%s%s%d", preFix, string(colAxis[j%colAxisLen]), i+1)
				iCell := &ICell{
					Axis:  axis,
					Value: "",
					Style: &Style{},
				}
				styleID, err := file.GetCellStyle(sheet1, axis)
				if err == nil {
					iCell.Style = d.GetStyle(file, styleID)
				}
				// 判断是被合并单元格
				if mgr, ok := d.mergeMG.HideCells[iCell.Axis]; ok {
					iCell.Hide = true
					mgr.MrCells = append(mgr.MrCells, iCell)
				}
				// 判断是被合并单元格@Y
				if mgr, ok := d.mergeMG.HideCellsY[iCell.Axis]; ok {
					iCell.Hide = true
					mgr.MrCellsY = append(mgr.MrCellsY, iCell)
				}
				rows[i] = append(rows[i], iCell)
			}
		}
	}
	return
}

func (d *Ex2Img) doMarge(cell *ICell) {
	// 判断是被合并单元格 执行合并
	if mgr, ok := d.mergeMG.MainCells[cell.Axis]; ok {
		switch mgr.TypeX {
		case true:
			w := cell.Width
			for _, mrCell := range mgr.MrCells {
				w += mrCell.Width
			}
			cell.Width = w
			if len(mgr.MrCellsY) > 0 {
				hRow := map[string]int{}
				for _, iCell := range mgr.MrCellsY {
					hRow[iCell.Axis[1:]] = iCell.Height
				}
				for _, i := range hRow {
					cell.Height += i
				}
			}
		case false:
			h := cell.Height
			for _, mrCell := range mgr.MrCells {
				h += mrCell.Height
			}
			cell.Height = h
		}
		cell.MrMain = true
	}
}

func (d *Ex2Img) draw(rows [][]*ICell) *image.RGBA {
	bg := image.White
	rgba := image.NewRGBA(image.Rect(0, 0, d.dWidth, d.dHeight))
	draw.Draw(rgba, rgba.Bounds(), bg, image.Point{}, draw.Src)
	var startY = 0
	for _, row := range rows {
		startY = d.drawRow(rgba, row, startY)
	}
	dst := image.NewRGBA(image.Rect(0, 0, d.dWidth, d.dHeight))
	draw.Draw(dst, image.Rect(0, 0, d.dWidth, d.dHeight), rgba, image.Point{1, 2}, draw.Src)
	return dst
}

func (d *Ex2Img) drawCell(cell *ICell) (*image.RGBA, int) {
	// 执行合并单元格
	d.doMarge(cell)
	fg, bg := cell.getFontColor(), cell.getBgColor()
	rgba := image.NewRGBA(image.Rect(0, 0, cell.Width*2, cell.Height))
	draw.Draw(rgba, rgba.Bounds(), image.NewUniform(bg), image.Point{}, draw.Src)
	c := freetype.NewContext()
	c.SetDPI(144)
	c.SetFont(cell.getFontTT())
	c.SetFontSize(float64(cell.getSize()))
	c.SetClip(rgba.Bounds())
	c.SetDst(rgba)
	c.SetSrc(image.NewUniform(fg))
	c.SetHinting(font.HintingFull)
	x, y := cell.Width, cell.getBeginPY()
	// draw
	pt := freetype.Pt(x, y)
	nPoint, err := c.DrawString(cell.Value, pt)
	if err != nil {
		return rgba, 0
	}
	// 实际宽
	trueWidth := nPoint.X.Ceil() - x
	// 下滑线
	if cell.Style.Font.Underline {
		ly := y + 1
		lx := x + trueWidth
		d.drawLine(rgba, x, ly, lx, ly, color.Black)
	}
	// 删除线
	if cell.Style.Font.Strike {
		ly := y - cell.getSize()/2
		lx := x + trueWidth
		d.drawLine(rgba, x, ly, lx, ly, color.Black)
	}
	// 更好的横对齐
	sx := cell.Width
	switch cell.Style.Alignment.Horizontal {
	case "center":
		sx = x - (cell.Width/2 - trueWidth/2)
	case "left":
		sx = sx - 2
	case "right":
		sx = x - (cell.Width - trueWidth)
	}
	// 画边框
	if cell.Style.Border.Left != "" {
		ruler := cell.getBorderLeftColor()
		d.drawLine(rgba, sx+1, 1, sx+1, cell.Height, ruler)
	}
	if cell.Style.Border.Top != "" {
		ruler := cell.getBorderTopColor()
		d.drawLine(rgba, sx+1, 1, sx+cell.Width, 1, ruler)
	}
	if cell.Style.Border.Right != "" {
		ruler := cell.getBorderRightColor()
		d.drawLine(rgba, sx+cell.Width-1, 1, sx+cell.Width-1, cell.Height, ruler)
	}
	if cell.Style.Border.Bottom != "" {
		ruler := cell.getBorderBottomColor()
		d.drawLine(rgba, sx+1, cell.Height-1, sx+cell.Width, cell.Height-1, ruler)
	}
	return rgba, sx
}

func (d *Ex2Img) drawLine(dst *image.RGBA, x, y, x2, y2 int, ruler color.Color) {
	for i := x; i <= x2; i++ {
		dst.Set(i, y, ruler)
	}
	for i := y; i <= y2; i++ {
		dst.Set(x, i, ruler)
	}
}

func (d *Ex2Img) drawRow(dst *image.RGBA, row []*ICell, startY int) (nextStartY int) {
	x, y, offsetY := 0, startY+1, 0
	cellY := row[0].Height
	for _, cell := range row {
		if cell.Hide {
			if mgr, ok := d.mergeMG.HideCells[cell.Axis]; ok {
				if !mgr.TypeX {
					x += cell.Width
				}
			}
			if _, ok := d.mergeMG.HideCellsY[cell.Axis]; ok {
				x += cell.Width
			}
			continue
		}
		ira, sx := d.drawCell(cell)
		if !cell.MrMain {
			offsetY = cell.Height
		}
		draw.Draw(dst, image.Rect(x, y, cell.Width+x, cell.Height+y), ira, image.Point{sx, 0}, draw.Src)
		x += cell.Width
	}
	if offsetY == 0 {
		offsetY = cellY
	}
	return startY + offsetY
}

func (d *Ex2Img) save(filename string, rgba *image.RGBA) error {
	// Save that RGBA image to disk.
	outFile, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer outFile.Close()
	b := bufio.NewWriter(outFile)
	err = png.Encode(b, rgba)
	if err != nil {
		return err
	}
	err = b.Flush()
	if err != nil {
		return err
	}
	return nil
}
