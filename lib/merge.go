package lib

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

type IMerge struct {
	Axis     string
	Value    string
	TypeX    bool
	MrCells  []*ICell
	MrCellsY []*ICell
}

type MergeMG struct {
	MainCells  map[string]*IMerge
	HideCells  map[string]*IMerge
	HideCellsY map[string]*IMerge
}

func NewMergeMG(mergeCells []excelize.MergeCell) *MergeMG {
	mg := &MergeMG{
		MainCells:  map[string]*IMerge{},
		HideCells:  map[string]*IMerge{},
		HideCellsY: map[string]*IMerge{},
	}
	for _, cell := range mergeCells {
		sa := cell.GetStartAxis()
		ea := cell.GetEndAxis()
		imr := &IMerge{
			Axis:     sa,
			Value:    cell.GetCellValue(),
			TypeX:    false,
			MrCells:  make([]*ICell, 0),
			MrCellsY: make([]*ICell, 0),
		}
		mg.MainCells[sa] = imr
		// 判断是横还是纵
		if sa[:1] == ea[:1] { // 纵
			ss, es := sa[1:], ea[1:]
			si, _ := strconv.Atoi(ss)
			ei, _ := strconv.Atoi(es)
			for i := si + 1; i <= ei; i++ {
				axis := fmt.Sprintf("%s%d", sa[:1], i)
				mg.HideCells[axis] = imr
			}
		} else { // 横
			imr.TypeX = true
			stl, etl := []rune(sa[:1]), []rune(ea[:1])
			for i := stl[0] + 1; i <= etl[0]; i++ {
				ac := []rune{i}
				axis := fmt.Sprintf("%s%s", string(ac), sa[1:])
				mg.HideCells[axis] = imr
			}
			// 先横后纵 多行合并
			if sa[1:] != ea[1:] {
				ss, es := sa[1:], ea[1:]
				si, _ := strconv.Atoi(ss)
				ei, _ := strconv.Atoi(es)
				for i := si + 1; i <= ei; i++ {
					for y := stl[0]; y <= etl[0]; y++ {
						ac := []rune{y}
						axis := fmt.Sprintf("%s%d", string(ac), i)
						mg.HideCellsY[axis] = imr
					}
				}
			}
		}
	}
	return mg
}
