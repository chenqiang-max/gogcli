package cmd

import (
	"encoding/json"
	"fmt"
)

// JSONPresentationSpec is the top-level JSON input structure.
type JSONPresentationSpec struct {
	Presentation JSONPresentation `json:"presentation"`
}

// JSONPresentation describes the presentation to create.
type JSONPresentation struct {
	Title      string          `json:"title"`
	TemplateID string          `json:"templateId,omitempty"`
	Slides     []JSONSlideSpec `json:"slides"`
}

// JSONSlideSpec describes a single slide.
type JSONSlideSpec struct {
	Layout         string                     `json:"layout"`
	Content        map[string]JSONContentValue `json:"content,omitempty"`
	Notes          string                     `json:"notes,omitempty"`
	CustomElements []JSONCustomElement         `json:"customElements,omitempty"`
}

// JSONContentValue holds a placeholder content value which can be either a
// plain string or a typed object (text, table, image).
type JSONContentValue struct {
	Type string // "text", "table", "image"

	// Text content
	Text string

	// Table content
	Rows      int        `json:"rows,omitempty"`
	Columns   int        `json:"columns,omitempty"`
	Data      [][]string `json:"data,omitempty"`
	HeaderRow bool       `json:"headerRow,omitempty"`

	// Image content
	URL string `json:"url,omitempty"`
}

// UnmarshalJSON handles both plain strings and typed objects.
func (v *JSONContentValue) UnmarshalJSON(data []byte) error {
	// Try plain string first
	var s string
	if err := json.Unmarshal(data, &s); err == nil {
		v.Type = "text"
		v.Text = s
		return nil
	}

	// Try typed object
	type rawContent struct {
		Type      string          `json:"type"`
		Text      string          `json:"text,omitempty"`
		Rows      int             `json:"rows,omitempty"`
		Columns   int             `json:"columns,omitempty"`
		Data      [][]interface{} `json:"data,omitempty"`
		HeaderRow bool            `json:"headerRow,omitempty"`
		URL       string          `json:"url,omitempty"`
	}

	var rc rawContent
	if err := json.Unmarshal(data, &rc); err != nil {
		return fmt.Errorf("content value must be a string or object: %w", err)
	}

	if rc.Type == "" {
		rc.Type = "text"
	}

	v.Type = rc.Type
	v.Text = rc.Text
	v.Rows = rc.Rows
	v.Columns = rc.Columns
	v.HeaderRow = rc.HeaderRow
	v.URL = rc.URL

	// Convert data cells to strings
	if len(rc.Data) > 0 {
		v.Data = make([][]string, len(rc.Data))
		for i, row := range rc.Data {
			v.Data[i] = make([]string, len(row))
			for j, cell := range row {
				if cell != nil {
					v.Data[i][j] = fmt.Sprintf("%v", cell)
				}
			}
		}
	}

	return nil
}

// JSONCustomElement represents a free-form element on a slide.
type JSONCustomElement struct {
	Type         string              `json:"type"` // "shape", "image", "line", "table"
	ID           string              `json:"id,omitempty"`
	ShapeType    string              `json:"shapeType,omitempty"`
	Text         string              `json:"text,omitempty"`
	URL          string              `json:"url,omitempty"`
	LineCategory string              `json:"lineCategory,omitempty"`
	Position     JSONPosition        `json:"position,omitempty"`
	Size         JSONSize            `json:"size,omitempty"`
	Style        *JSONElementStyle   `json:"style,omitempty"`

	// Table fields
	Rows      int             `json:"rows,omitempty"`
	Columns   int             `json:"columns,omitempty"`
	Data      [][]interface{} `json:"data,omitempty"`
	HeaderRow bool            `json:"headerRow,omitempty"`

	// Line connections
	StartConnection *JSONLineConnection `json:"startConnection,omitempty"`
	EndConnection   *JSONLineConnection `json:"endConnection,omitempty"`
}

// JSONPosition describes element positioning.
type JSONPosition struct {
	X    float64 `json:"x"`
	Y    float64 `json:"y"`
	Unit string  `json:"unit,omitempty"`
}

// JSONSize describes element dimensions.
type JSONSize struct {
	Width  float64 `json:"width"`
	Height float64 `json:"height"`
	Unit   string  `json:"unit,omitempty"`
}

// JSONElementStyle describes styling for custom elements.
type JSONElementStyle struct {
	BackgroundColor *JSONColor     `json:"backgroundColor,omitempty"`
	TextStyle       *JSONTextStyle `json:"textStyle,omitempty"`
	LineStyle       *JSONLineStyle `json:"lineStyle,omitempty"`
}

// JSONColor specifies a color via theme name or RGB tuple.
type JSONColor struct {
	Theme string    `json:"theme,omitempty"`
	RGB   []float64 `json:"rgb,omitempty"`
	Alpha float64   `json:"alpha,omitempty"`
}

// JSONTextStyle describes text styling.
type JSONTextStyle struct {
	FontSize   float64    `json:"fontSize,omitempty"`
	Bold       *bool      `json:"bold,omitempty"`
	Italic     *bool      `json:"italic,omitempty"`
	FontFamily string     `json:"fontFamily,omitempty"`
	Color      *JSONColor `json:"color,omitempty"`
}

// JSONLineStyle describes line styling.
type JSONLineStyle struct {
	Weight     float64 `json:"weight,omitempty"`
	EndArrow   string  `json:"endArrow,omitempty"`
	StartArrow string  `json:"startArrow,omitempty"`
}

// JSONLineConnection describes a line connection to another element.
type JSONLineConnection struct {
	ElementID string `json:"elementId"`
	Site      int64  `json:"site,omitempty"`
}
