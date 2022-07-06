package lib

type Border struct {
	Left        string
	LeftColor   string
	Right       string
	RightColor  string
	Top         string
	TopColor    string
	Bottom      string
	BottomColor string
}

type Style struct {
	Border         Border
	Fill           Fill
	Font           Font
	ApplyBorder    bool
	ApplyFill      bool
	ApplyFont      bool
	ApplyAlignment bool
	Alignment      Alignment
}

type Fill struct {
	PatternType string
	BgColor     string
	FgColor     string
	Tint        float64
}

type Font struct {
	Size      float64
	Name      string
	Color     string
	Bold      bool
	Italic    bool
	Underline bool
	Strike    bool
}

type Alignment struct {
	Horizontal   string
	Indent       int
	ShrinkToFit  bool
	TextRotation int
	Vertical     string
	WrapText     bool
}
