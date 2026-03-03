package cmd

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"regexp"
	"strings"
	"time"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/slides/v1"

	"github.com/steipete/gogcli/internal/outfmt"
	"github.com/steipete/gogcli/internal/ui"
)

// SlidesCreateFromJSONCmd creates a presentation from a JSON specification.
type SlidesCreateFromJSONCmd struct {
	Title       string `arg:"" name:"title" help:"Presentation title (overrides JSON spec title)" optional:""`
	Content     string `name:"content" help:"Inline JSON spec string (alternative to --content-file)"`
	ContentFile string `name:"content-file" help:"Path to a JSON spec file (alternative to --content)"`
	Parent      string `name:"parent" help:"Destination Google Drive folder ID to move the created presentation into"`
	Template    string `name:"template" help:"Template presentation ID to copy from (overrides JSON templateId); the template's slides are replaced by the spec"`
	Debug       bool   `name:"debug" help:"Show debug output including request counts and placeholder mappings"`
}

// layoutPlaceholders defines the expected placeholders per predefined layout.
var layoutPlaceholders = map[string][]string{
	"TITLE_SLIDE":                   {"CENTERED_TITLE", "SUBTITLE"},
	"SECTION_HEADER":                {"TITLE"},
	"TITLE_AND_BODY":                {"TITLE", "BODY"},
	"TITLE_AND_TWO_COLUMNS":         {"TITLE", "BODY_0", "BODY_1"},
	"TITLE_ONLY":                    {"TITLE"},
	"ONE_COLUMN_TEXT":               {"TITLE", "BODY"},
	"MAIN_POINT":                    {"TITLE"},
	"SECTION_TITLE_AND_DESCRIPTION": {"TITLE", "SUBTITLE", "BODY"},
	"CAPTION":                       {"BODY"},
	"BIG_NUMBER":                    {"TITLE", "BODY"},
	"BLANK":                         {},
}

func (c *SlidesCreateFromJSONCmd) Run(ctx context.Context, flags *RootFlags) error {
	u := ui.FromContext(ctx)
	account, err := getAccountForDrive(flags)
	if err != nil {
		return err
	}

	// Parse JSON spec
	spec, err := c.parseSpec()
	if err != nil {
		return err
	}

	// Resolve title
	title := strings.TrimSpace(c.Title)
	if title == "" {
		title = strings.TrimSpace(spec.Presentation.Title)
	}
	if title == "" {
		return usage("empty title: provide as argument or in JSON spec")
	}

	if len(spec.Presentation.Slides) == 0 {
		return usage("JSON spec must contain at least one slide")
	}

	templateID := strings.TrimSpace(c.Template)
	if templateID == "" {
		templateID = strings.TrimSpace(spec.Presentation.TemplateID)
	}

	if c.Debug {
		debugSlides = true
	}

	// Dry-run: print what would happen and exit before any API calls
	if err := dryRunExit(ctx, flags, "slides.create-from-json", map[string]any{
		"title":      title,
		"templateId": templateID,
		"parent":     c.Parent,
		"slides":     len(spec.Presentation.Slides),
	}); err != nil {
		return err
	}

	// Create services
	slidesSvc, err := newSlidesService(ctx, account)
	if err != nil {
		return err
	}
	driveSvc, err := newDriveService(ctx, account)
	if err != nil {
		return err
	}

	// Step 1: Create or copy presentation
	var presID string
	if templateID != "" {
		u.Err().Printf("Copying template %s as '%s'...", templateID, title)
		f := &drive.File{Name: title}
		created, copyErr := driveSvc.Files.Copy(templateID, f).
			SupportsAllDrives(true).
			Fields("id").
			Context(ctx).
			Do()
		if copyErr != nil {
			return fmt.Errorf("failed to copy template: %w", copyErr)
		}
		presID = created.Id
	} else {
		u.Err().Printf("Creating presentation '%s'...", title)
		pres, createErr := slidesSvc.Presentations.Create(&slides.Presentation{
			Title: title,
		}).Context(ctx).Do()
		if createErr != nil {
			return fmt.Errorf("failed to create presentation: %w", createErr)
		}
		presID = pres.PresentationId
	}

	// Step 2: Read presentation to get layouts
	pres, err := slidesSvc.Presentations.Get(presID).Context(ctx).Do()
	if err != nil {
		return fmt.Errorf("failed to get presentation: %w", err)
	}

	layoutMap := buildLayoutMap(pres)
	existingSlideIDs := make([]string, 0, len(pres.Slides))
	for _, s := range pres.Slides {
		existingSlideIDs = append(existingSlideIDs, s.ObjectId)
	}

	if debugSlides {
		fmt.Fprintf(os.Stderr, "[DEBUG] Found %d layouts, %d existing slides\n",
			len(layoutMap), len(existingSlideIDs))
	}

	// Step 3: Pass 1 — Create slides + populate content
	var pass1Reqs []*slides.Request
	type slideNotes struct {
		slideID string
		notes   string
	}
	var notesEntries []slideNotes

	for i, slideSpec := range spec.Presentation.Slides {
		slideID := fmt.Sprintf("json_slide_%d_%d", i, time.Now().UnixNano()%100000)
		layoutName := slideSpec.Layout

		// Build CreateSlide request
		createReq, phMap := buildCreateSlideRequest(slideID, layoutName, layoutMap, i)
		pass1Reqs = append(pass1Reqs, createReq)

		if debugSlides {
			fmt.Fprintf(os.Stderr, "[DEBUG] Slide %d: layout=%s, id=%s, placeholders=%v\n",
				i+1, layoutName, slideID, phMap)
		}

		// Track notes
		if slideSpec.Notes != "" {
			notesEntries = append(notesEntries, slideNotes{slideID: slideID, notes: slideSpec.Notes})
		}

		// Fill placeholder content
		for phKey, content := range slideSpec.Content {
			phObjID, ok := phMap[phKey]
			if !ok {
				u.Err().Printf("  [WARNING] Placeholder %s not available in layout %s. Skipping.", phKey, layoutName)
				continue
			}

			switch content.Type {
			case "text", "":
				reqs := buildInsertTextRequests(phObjID, content.Text)
				pass1Reqs = append(pass1Reqs, reqs...)

			case "table":
				// Delete placeholder and create table in its place
				pass1Reqs = append(pass1Reqs, &slides.Request{
					DeleteObject: &slides.DeleteObjectRequest{ObjectId: phObjID},
				})
				reqs := buildPlaceholderTableRequests(slideID, content)
				pass1Reqs = append(pass1Reqs, reqs...)

			case "image":
				// Delete placeholder and create image in its place
				pass1Reqs = append(pass1Reqs, &slides.Request{
					DeleteObject: &slides.DeleteObjectRequest{ObjectId: phObjID},
				})
				reqs := buildPlaceholderImageRequests(slideID, content)
				pass1Reqs = append(pass1Reqs, reqs...)
			}
		}

		// Custom elements
		for _, elem := range slideSpec.CustomElements {
			elemType := elem.Type
			if elemType == "" {
				elemType = "shape"
			}
			switch elemType {
			case "shape":
				pass1Reqs = append(pass1Reqs, buildCustomShapeRequests(slideID, elem)...)
			case "image":
				pass1Reqs = append(pass1Reqs, buildCustomImageRequests(slideID, elem)...)
			case "line":
				pass1Reqs = append(pass1Reqs, buildCustomLineRequests(slideID, elem)...)
			case "table":
				pass1Reqs = append(pass1Reqs, buildCustomTableRequests(slideID, elem)...)
			}
		}
	}

	// Delete existing slides (template slides or the default blank)
	for _, sid := range existingSlideIDs {
		pass1Reqs = append(pass1Reqs, &slides.Request{
			DeleteObject: &slides.DeleteObjectRequest{ObjectId: sid},
		})
	}

	if debugSlides {
		fmt.Fprintf(os.Stderr, "[DEBUG] Pass 1: sending %d requests\n", len(pass1Reqs))
	}

	if len(pass1Reqs) > 0 {
		_, err = slidesSvc.Presentations.BatchUpdate(presID, &slides.BatchUpdatePresentationRequest{
			Requests: pass1Reqs,
		}).Context(ctx).Do()
		if err != nil {
			return fmt.Errorf("pass 1 batch update failed: %w", err)
		}
	}

	// Step 4: Pass 2 — Speaker notes
	if len(notesEntries) > 0 {
		if debugSlides {
			fmt.Fprintf(os.Stderr, "[DEBUG] Pass 2: adding notes to %d slides\n", len(notesEntries))
		}

		pres, err = slidesSvc.Presentations.Get(presID).Context(ctx).Do()
		if err != nil {
			return fmt.Errorf("failed to re-fetch presentation for notes: %w", err)
		}

		slideMap := make(map[string]*slides.Page, len(pres.Slides))
		for _, s := range pres.Slides {
			slideMap[s.ObjectId] = s
		}

		var pass2Reqs []*slides.Request
		for _, entry := range notesEntries {
			slide := slideMap[entry.slideID]
			if slide == nil {
				u.Err().Printf("  [WARNING] Slide %s not found for notes. Skipping.", entry.slideID)
				continue
			}

			notesObjID := findNotesPlaceholder(slide)
			if notesObjID == "" {
				u.Err().Printf("  [WARNING] No notes placeholder found for slide %s.", entry.slideID)
				continue
			}

			pass2Reqs = append(pass2Reqs, buildInsertTextRequests(notesObjID, entry.notes)...)
		}

		if len(pass2Reqs) > 0 {
			_, err = slidesSvc.Presentations.BatchUpdate(presID, &slides.BatchUpdatePresentationRequest{
				Requests: pass2Reqs,
			}).Context(ctx).Do()
			if err != nil {
				return fmt.Errorf("pass 2 (notes) batch update failed: %w", err)
			}
		}
	}

	// Step 5: Move to parent folder if specified
	if c.Parent != "" {
		_, err = driveSvc.Files.Update(presID, &drive.File{}).
			AddParents(c.Parent).
			SupportsAllDrives(true).
			Context(ctx).
			Do()
		if err != nil {
			return fmt.Errorf("failed to move presentation to folder: %w", err)
		}
	}

	// Step 6: Fetch file details for output
	file, err := driveSvc.Files.Get(presID).
		Fields("id, name, webViewLink").
		SupportsAllDrives(true).
		Context(ctx).
		Do()
	if err != nil {
		return fmt.Errorf("failed to get file details: %w", err)
	}

	if outfmt.IsJSON(ctx) {
		return outfmt.WriteJSON(ctx, os.Stdout, map[string]any{
			"presentationId": presID,
			"file":           file,
		})
	}

	u.Out().Printf("Created presentation with %d slides", len(spec.Presentation.Slides))
	u.Out().Printf("id\t%s", presID)
	u.Out().Printf("name\t%s", file.Name)
	if file.WebViewLink != "" {
		u.Out().Printf("link\t%s", file.WebViewLink)
	}
	return nil
}

// parseSpec reads and parses the JSON specification from --content or --content-file.
func (c *SlidesCreateFromJSONCmd) parseSpec() (*JSONPresentationSpec, error) {
	var raw string
	switch {
	case c.ContentFile != "":
		data, err := os.ReadFile(c.ContentFile)
		if err != nil {
			return nil, fmt.Errorf("failed to read content file: %w", err)
		}
		raw = string(data)
	case c.Content != "":
		raw = c.Content
	default:
		return nil, usage("either --content or --content-file is required")
	}

	var spec JSONPresentationSpec
	if err := json.Unmarshal([]byte(raw), &spec); err != nil {
		return nil, fmt.Errorf("failed to parse JSON spec: %w", err)
	}

	return &spec, nil
}

// buildLayoutMap builds a map from layout display name (and normalized variants)
// to layout object ID from the presentation's layouts.
func buildLayoutMap(pres *slides.Presentation) map[string]string {
	m := make(map[string]string)
	for _, layout := range pres.Layouts {
		name := layout.LayoutProperties.DisplayName
		m[name] = layout.ObjectId
		// Also store uppercase/underscore-normalized form
		normalized := strings.ToUpper(strings.ReplaceAll(name, " ", "_"))
		m[normalized] = layout.ObjectId
	}
	return m
}

// buildCreateSlideRequest builds a CreateSlide request with placeholder mappings.
// Returns the request and a map of placeholder key → assigned object ID.
func buildCreateSlideRequest(slideID, layoutName string, layoutMap map[string]string, index int) (*slides.Request, map[string]string) {
	phMap := make(map[string]string)

	req := &slides.Request{
		CreateSlide: &slides.CreateSlideRequest{
			ObjectId: slideID,
		},
	}

	// Try layout map first (template/custom layouts)
	if objID, ok := layoutMap[layoutName]; ok {
		req.CreateSlide.SlideLayoutReference = &slides.LayoutReference{
			LayoutId: objID,
		}
	} else {
		// Normalize to predefined layout enum
		normalized := normalizePredefinedLayout(layoutName)
		req.CreateSlide.SlideLayoutReference = &slides.LayoutReference{
			PredefinedLayout: normalized,
		}
	}

	// Build placeholder ID mappings
	placeholders := resolveLayoutPlaceholders(layoutName)
	var mappings []*slides.LayoutPlaceholderIdMapping
	for _, ph := range placeholders {
		phType, phIndex := parsePlaceholderType(ph)
		objID := fmt.Sprintf("ph_%s_%d_%d", strings.ToLower(ph), index, time.Now().UnixNano()%100000)
		phMap[ph] = objID

		mapping := &slides.LayoutPlaceholderIdMapping{
			ObjectId: objID,
			LayoutPlaceholder: &slides.Placeholder{
				Type: phType,
			},
		}
		if phIndex >= 0 {
			mapping.LayoutPlaceholder.Index = int64(phIndex)
			mapping.LayoutPlaceholder.ForceSendFields = []string{"Index"}
		}
		mappings = append(mappings, mapping)
	}

	if len(mappings) > 0 {
		req.CreateSlide.PlaceholderIdMappings = mappings
	}

	return req, phMap
}

// normalizePredefinedLayout normalizes a JSON layout name to the Google Slides
// API PredefinedLayout enum value. The JSON spec uses friendly names like
// "TITLE_SLIDE" and "CAPTION" while the API uses "TITLE" and "CAPTION_ONLY".
func normalizePredefinedLayout(name string) string {
	normalized := strings.ToUpper(strings.ReplaceAll(strings.TrimSpace(name), " ", "_"))
	switch normalized {
	case "TITLE_SLIDE":
		return "TITLE"
	case "CAPTION":
		return "CAPTION_ONLY"
	default:
		return normalized
	}
}

// resolveLayoutPlaceholders returns the expected placeholder keys for a layout.
func resolveLayoutPlaceholders(layoutName string) []string {
	normalized := strings.ToUpper(strings.ReplaceAll(strings.TrimSpace(layoutName), " ", "_"))
	if phs, ok := layoutPlaceholders[normalized]; ok {
		return phs
	}
	return nil
}

// parsePlaceholderType converts a placeholder key like "BODY_0" to (type, index).
func parsePlaceholderType(key string) (string, int) {
	// Handle indexed placeholders like BODY_0, BODY_1
	if idx := strings.LastIndex(key, "_"); idx > 0 {
		suffix := key[idx+1:]
		if len(suffix) == 1 && suffix[0] >= '0' && suffix[0] <= '9' {
			return key[:idx], int(suffix[0] - '0')
		}
	}
	return key, -1
}

// findNotesPlaceholder finds the speaker notes placeholder ID for a slide.
func findNotesPlaceholder(slide *slides.Page) string {
	if slide.SlideProperties == nil || slide.SlideProperties.NotesPage == nil {
		return ""
	}
	np := slide.SlideProperties.NotesPage

	// Prefer the direct property
	if np.NotesProperties != nil && np.NotesProperties.SpeakerNotesObjectId != "" {
		return np.NotesProperties.SpeakerNotesObjectId
	}

	// Fallback: search page elements for BODY placeholder
	for _, el := range np.PageElements {
		if el.Shape != nil && el.Shape.Placeholder != nil &&
			el.Shape.Placeholder.Type == placeholderTypeBody {
			return el.ObjectId
		}
	}
	return ""
}

// ---------------------------------------------------------------------------
// Markdown text parsing and formatting
// ---------------------------------------------------------------------------

type textAnnotation struct {
	start int
	end   int
	style string // "bold", "italic", "strikethrough", "underline"
	url   string // for "link" style
}

// parseMarkdownText parses simple markdown formatting into plain text + annotations.
func parseMarkdownText(input string) (string, []textAnnotation) {
	var plain strings.Builder
	var annotations []textAnnotation

	type patternDef struct {
		re    *regexp.Regexp
		style string
	}

	// Process in order: bold, italic, strikethrough, underline, links
	// We'll do a multi-pass approach on the input string
	patterns := []patternDef{
		{regexp.MustCompile(`\*\*(.+?)\*\*`), "bold"},
		{regexp.MustCompile(`\*(.+?)\*`), "italic"},
		{regexp.MustCompile(`~~(.+?)~~`), "strikethrough"},
		{regexp.MustCompile(`__(.+?)__`), "underline"},
	}
	linkPattern := regexp.MustCompile(`\[([^\]]+)\]\(([^)]+)\)`)

	// First pass: replace links
	type replacement struct {
		origStart, origEnd int
		text               string
		style              string
		url                string
	}

	var replacements []replacement

	// Find links
	for _, match := range linkPattern.FindAllStringSubmatchIndex(input, -1) {
		replacements = append(replacements, replacement{
			origStart: match[0],
			origEnd:   match[1],
			text:      input[match[2]:match[3]],
			style:     "link",
			url:       input[match[4]:match[5]],
		})
	}

	// Build intermediate text with links replaced
	var intermediate strings.Builder
	lastEnd := 0
	type linkAnnotation struct {
		start int
		text  string
		url   string
	}
	var linkAnnotations []linkAnnotation

	for _, r := range replacements {
		intermediate.WriteString(input[lastEnd:r.origStart])
		linkStart := intermediate.Len()
		intermediate.WriteString(r.text)
		linkAnnotations = append(linkAnnotations, linkAnnotation{
			start: linkStart,
			text:  r.text,
			url:   r.url,
		})
		lastEnd = r.origEnd
	}
	intermediate.WriteString(input[lastEnd:])
	text := intermediate.String()

	// Process formatting patterns
	type fmtAnnotation struct {
		start, end int
		style      string
	}
	var fmtAnnotations []fmtAnnotation

	// We need to process the text removing markers while tracking positions
	// Do a simple sequential approach
	result := text
	offset := 0

	for _, p := range patterns {
		matches := p.re.FindAllStringSubmatchIndex(result, -1)
		if len(matches) == 0 {
			continue
		}

		var newResult strings.Builder
		lastIdx := 0
		localOffset := 0
		for _, match := range matches {
			newResult.WriteString(result[lastIdx:match[0]])
			contentStart := newResult.Len()
			content := result[match[2]:match[3]]
			newResult.WriteString(content)
			contentEnd := newResult.Len()
			fmtAnnotations = append(fmtAnnotations, fmtAnnotation{
				start: contentStart + offset,
				end:   contentEnd + offset,
				style: p.style,
			})
			localOffset += (match[1] - match[0]) - len(content)
			lastIdx = match[1]
		}
		newResult.WriteString(result[lastIdx:])
		offset += 0 // offset tracking is handled differently
		result = newResult.String()
	}

	plain.WriteString(result)

	// Convert to textAnnotation slice
	for _, a := range fmtAnnotations {
		annotations = append(annotations, textAnnotation{
			start: a.start,
			end:   a.end,
			style: a.style,
		})
	}

	// Add link annotations (adjust for formatting changes)
	for _, la := range linkAnnotations {
		// Find the link text in the final plain text
		idx := strings.Index(plain.String(), la.text)
		if idx >= 0 {
			annotations = append(annotations, textAnnotation{
				start: idx,
				end:   idx + len(la.text),
				style: "link",
				url:   la.url,
			})
		}
	}

	return plain.String(), annotations
}

// normalizeNewlines converts paragraph separators (\n\n) into Google Slides
// paragraph breaks (\n) and single newlines (\n) into soft line breaks (\v,
// equivalent to Shift+Enter in the UI).  Annotation offsets are adjusted to
// account for the shorter output.
func normalizeNewlines(text string, annotations []textAnnotation) (string, []textAnnotation) {
	// Fast path: nothing to do if there are no newlines.
	if !strings.Contains(text, "\n") {
		return text, annotations
	}

	var out strings.Builder
	out.Grow(len(text))

	// offsetMap[i] = position in output that corresponds to position i in input
	offsetMap := make([]int, len(text)+1)
	inPos := 0
	outPos := 0

	for inPos < len(text) {
		offsetMap[inPos] = outPos
		if inPos+1 < len(text) && text[inPos] == '\n' && text[inPos+1] == '\n' {
			// \n\n → single \n (paragraph break)
			out.WriteByte('\n')
			offsetMap[inPos+1] = outPos // second \n maps to same output position
			inPos += 2
			outPos++
		} else if text[inPos] == '\n' {
			// single \n → \v (soft line break)
			out.WriteByte('\v')
			inPos++
			outPos++
		} else {
			out.WriteByte(text[inPos])
			inPos++
			outPos++
		}
	}
	offsetMap[inPos] = outPos

	// Adjust annotations
	adjusted := make([]textAnnotation, len(annotations))
	for i, a := range annotations {
		adjusted[i] = textAnnotation{
			start: offsetMap[clamp(a.start, 0, len(text))],
			end:   offsetMap[clamp(a.end, 0, len(text))],
			style: a.style,
			url:   a.url,
		}
	}

	return out.String(), adjusted
}

func clamp(v, lo, hi int) int {
	if v < lo {
		return lo
	}
	if v > hi {
		return hi
	}
	return v
}

// buildInsertTextRequests builds InsertText + UpdateTextStyle requests for markdown text.
// Newline handling: \n\n → paragraph break (\n), single \n → soft line break (\v / Shift+Enter).
func buildInsertTextRequests(objectID, text string) []*slides.Request {
	plainText, annotations := parseMarkdownText(text)
	if plainText == "" {
		return nil
	}

	// Normalize newlines: \n\n → \n (paragraph break), single \n → \v (soft break)
	plainText, annotations = normalizeNewlines(plainText, annotations)

	var reqs []*slides.Request

	reqs = append(reqs, &slides.Request{
		InsertText: &slides.InsertTextRequest{
			ObjectId:       objectID,
			Text:           plainText,
			InsertionIndex: 0,
		},
	})

	for _, a := range annotations {
		var style *slides.TextStyle
		var fields string

		switch a.style {
		case "bold":
			style = &slides.TextStyle{Bold: true}
			fields = "bold"
		case "italic":
			style = &slides.TextStyle{Italic: true}
			fields = "italic"
		case "strikethrough":
			style = &slides.TextStyle{Strikethrough: true}
			fields = "strikethrough"
		case "underline":
			style = &slides.TextStyle{Underline: true}
			fields = "underline"
		case "link":
			style = &slides.TextStyle{
				Link: &slides.Link{Url: a.url},
			}
			fields = "link"
		default:
			continue
		}

		startIdx := int64(a.start)
		endIdx := int64(a.end)
		reqs = append(reqs, &slides.Request{
			UpdateTextStyle: &slides.UpdateTextStyleRequest{
				ObjectId: objectID,
				TextRange: &slides.Range{
					Type:       "FIXED_RANGE",
					StartIndex: &startIdx,
					EndIndex:   &endIdx,
				},
				Style:  style,
				Fields: fields,
			},
		})
	}

	return reqs
}

// ---------------------------------------------------------------------------
// Placeholder content builders
// ---------------------------------------------------------------------------

// buildPlaceholderTableRequests builds a table in place of a deleted placeholder.
func buildPlaceholderTableRequests(slideID string, content JSONContentValue) []*slides.Request {
	rows := content.Rows
	cols := content.Columns
	if rows <= 0 && len(content.Data) > 0 {
		rows = len(content.Data)
	}
	if cols <= 0 && len(content.Data) > 0 && len(content.Data[0]) > 0 {
		cols = len(content.Data[0])
	}
	if rows <= 0 {
		rows = 1
	}
	if cols <= 0 {
		cols = 1
	}

	tableID := fmt.Sprintf("tbl_%d", time.Now().UnixNano())

	var reqs []*slides.Request
	reqs = append(reqs, &slides.Request{
		CreateTable: &slides.CreateTableRequest{
			ObjectId: tableID,
			ElementProperties: &slides.PageElementProperties{
				PageObjectId: slideID,
			},
			Rows:    int64(rows),
			Columns: int64(cols),
		},
	})

	// Fill cells
	for rIdx, row := range content.Data {
		for cIdx, cellVal := range row {
			if cellVal == "" {
				continue
			}
			reqs = append(reqs, &slides.Request{
				InsertText: &slides.InsertTextRequest{
					ObjectId: tableID,
					CellLocation: &slides.TableCellLocation{
						RowIndex:    int64(rIdx),
						ColumnIndex: int64(cIdx),
					},
					Text:           cellVal,
					InsertionIndex: 0,
				},
			})
		}
	}

	// Bold header row
	if content.HeaderRow && len(content.Data) > 0 {
		for cIdx := range content.Data[0] {
			if content.Data[0][cIdx] == "" {
				continue
			}
			reqs = append(reqs, &slides.Request{
				UpdateTextStyle: &slides.UpdateTextStyleRequest{
					ObjectId: tableID,
					CellLocation: &slides.TableCellLocation{
						RowIndex:    0,
						ColumnIndex: int64(cIdx),
					},
					Style: &slides.TextStyle{Bold: true},
					TextRange: &slides.Range{
						Type: "ALL",
					},
					Fields: "bold",
				},
			})
		}
	}

	return reqs
}

// buildPlaceholderImageRequests builds an image in place of a deleted placeholder.
func buildPlaceholderImageRequests(slideID string, content JSONContentValue) []*slides.Request {
	imageID := fmt.Sprintf("img_%d", time.Now().UnixNano())
	return []*slides.Request{
		{
			CreateImage: &slides.CreateImageRequest{
				ObjectId: imageID,
				Url:      content.URL,
				ElementProperties: &slides.PageElementProperties{
					PageObjectId: slideID,
					Size: &slides.Size{
						Width:  &slides.Dimension{Magnitude: 400, Unit: "PT"},
						Height: &slides.Dimension{Magnitude: 300, Unit: "PT"},
					},
					Transform: &slides.AffineTransform{
						ScaleX:     1,
						ScaleY:     1,
						TranslateX: 72,
						TranslateY: 72,
						Unit:       "PT",
					},
				},
			},
		},
	}
}

// ---------------------------------------------------------------------------
// Custom element builders
// ---------------------------------------------------------------------------

func resolveUnit(pos JSONPosition) string {
	if pos.Unit != "" {
		return pos.Unit
	}
	return "PT"
}

func buildCustomShapeRequests(slideID string, elem JSONCustomElement) []*slides.Request {
	shapeID := ensureValidID(elem.ID, "shp")
	unit := resolveUnit(elem.Position)
	shapeType := elem.ShapeType
	if shapeType == "" {
		shapeType = "TEXT_BOX"
	}

	width := elem.Size.Width
	if width == 0 {
		width = 200
	}
	height := elem.Size.Height
	if height == 0 {
		height = 100
	}

	var reqs []*slides.Request
	reqs = append(reqs, &slides.Request{
		CreateShape: &slides.CreateShapeRequest{
			ObjectId:  shapeID,
			ShapeType: shapeType,
			ElementProperties: &slides.PageElementProperties{
				PageObjectId: slideID,
				Size: &slides.Size{
					Width:  &slides.Dimension{Magnitude: width, Unit: unit},
					Height: &slides.Dimension{Magnitude: height, Unit: unit},
				},
				Transform: &slides.AffineTransform{
					ScaleX:     1,
					ScaleY:     1,
					TranslateX: elem.Position.X,
					TranslateY: elem.Position.Y,
					Unit:       unit,
				},
			},
		},
	})

	if elem.Text != "" {
		reqs = append(reqs, buildInsertTextRequests(shapeID, elem.Text)...)
	}

	if elem.Style != nil {
		reqs = append(reqs, buildElementStyleRequests(shapeID, elem.Style)...)
	}

	return reqs
}

func buildCustomImageRequests(slideID string, elem JSONCustomElement) []*slides.Request {
	imageID := ensureValidID(elem.ID, "img")
	unit := resolveUnit(elem.Position)

	width := elem.Size.Width
	if width == 0 {
		width = 400
	}
	height := elem.Size.Height
	if height == 0 {
		height = 300
	}

	return []*slides.Request{
		{
			CreateImage: &slides.CreateImageRequest{
				ObjectId: imageID,
				Url:      elem.URL,
				ElementProperties: &slides.PageElementProperties{
					PageObjectId: slideID,
					Size: &slides.Size{
						Width:  &slides.Dimension{Magnitude: width, Unit: unit},
						Height: &slides.Dimension{Magnitude: height, Unit: unit},
					},
					Transform: &slides.AffineTransform{
						ScaleX:     1,
						ScaleY:     1,
						TranslateX: elem.Position.X,
						TranslateY: elem.Position.Y,
						Unit:       unit,
					},
				},
			},
		},
	}
}

func buildCustomLineRequests(slideID string, elem JSONCustomElement) []*slides.Request {
	lineID := ensureValidID(elem.ID, "ln")
	unit := resolveUnit(elem.Position)
	lineCategory := elem.LineCategory
	if lineCategory == "" {
		lineCategory = "STRAIGHT"
	}

	width := elem.Size.Width
	if width == 0 {
		width = 300
	}

	var reqs []*slides.Request
	reqs = append(reqs, &slides.Request{
		CreateLine: &slides.CreateLineRequest{
			ObjectId:     lineID,
			LineCategory: lineCategory,
			ElementProperties: &slides.PageElementProperties{
				PageObjectId: slideID,
				Size: &slides.Size{
					Width:  &slides.Dimension{Magnitude: width, Unit: unit},
					Height: &slides.Dimension{Magnitude: elem.Size.Height, Unit: unit},
				},
				Transform: &slides.AffineTransform{
					ScaleX:     1,
					ScaleY:     1,
					TranslateX: elem.Position.X,
					TranslateY: elem.Position.Y,
					Unit:       unit,
				},
			},
		},
	})

	// Line properties
	var lineProps slides.LineProperties
	var fields []string

	if elem.Style != nil && elem.Style.LineStyle != nil {
		ls := elem.Style.LineStyle
		if ls.Weight > 0 {
			lineProps.Weight = &slides.Dimension{Magnitude: ls.Weight, Unit: "PT"}
			fields = append(fields, "weight")
		}
		if ls.EndArrow != "" {
			lineProps.EndArrow = ls.EndArrow
			fields = append(fields, "endArrow")
		}
		if ls.StartArrow != "" {
			lineProps.StartArrow = ls.StartArrow
			fields = append(fields, "startArrow")
		}
	}

	if elem.StartConnection != nil {
		lineProps.StartConnection = &slides.LineConnection{
			ConnectedObjectId:   elem.StartConnection.ElementID,
			ConnectionSiteIndex: int64(elem.StartConnection.Site),
		}
		fields = append(fields, "startConnection")
	}
	if elem.EndConnection != nil {
		lineProps.EndConnection = &slides.LineConnection{
			ConnectedObjectId:   elem.EndConnection.ElementID,
			ConnectionSiteIndex: int64(elem.EndConnection.Site),
		}
		fields = append(fields, "endConnection")
	}

	if len(fields) > 0 {
		reqs = append(reqs, &slides.Request{
			UpdateLineProperties: &slides.UpdateLinePropertiesRequest{
				ObjectId:       lineID,
				LineProperties: &lineProps,
				Fields:         strings.Join(fields, ","),
			},
		})
	}

	return reqs
}

func buildCustomTableRequests(slideID string, elem JSONCustomElement) []*slides.Request {
	rows := elem.Rows
	cols := elem.Columns
	if rows <= 0 && len(elem.Data) > 0 {
		rows = len(elem.Data)
	}
	if cols <= 0 && len(elem.Data) > 0 && len(elem.Data[0]) > 0 {
		cols = len(elem.Data[0])
	}
	if rows <= 0 {
		rows = 1
	}
	if cols <= 0 {
		cols = 1
	}

	tableID := ensureValidID(elem.ID, "tbl")

	var reqs []*slides.Request
	reqs = append(reqs, &slides.Request{
		CreateTable: &slides.CreateTableRequest{
			ObjectId: tableID,
			ElementProperties: &slides.PageElementProperties{
				PageObjectId: slideID,
			},
			Rows:    int64(rows),
			Columns: int64(cols),
		},
	})

	for rIdx, row := range elem.Data {
		for cIdx, cell := range row {
			cellVal := fmt.Sprintf("%v", cell)
			if cellVal == "" || cellVal == "<nil>" {
				continue
			}
			reqs = append(reqs, &slides.Request{
				InsertText: &slides.InsertTextRequest{
					ObjectId: tableID,
					CellLocation: &slides.TableCellLocation{
						RowIndex:    int64(rIdx),
						ColumnIndex: int64(cIdx),
					},
					Text:           cellVal,
					InsertionIndex: 0,
				},
			})
		}
	}

	if elem.HeaderRow && len(elem.Data) > 0 {
		for cIdx, cell := range elem.Data[0] {
			cellVal := fmt.Sprintf("%v", cell)
			if cellVal == "" || cellVal == "<nil>" {
				continue
			}
			reqs = append(reqs, &slides.Request{
				UpdateTextStyle: &slides.UpdateTextStyleRequest{
					ObjectId: tableID,
					CellLocation: &slides.TableCellLocation{
						RowIndex:    0,
						ColumnIndex: int64(cIdx),
					},
					Style: &slides.TextStyle{Bold: true},
					TextRange: &slides.Range{
						Type: "ALL",
					},
					Fields: "bold",
				},
			})
		}
	}

	return reqs
}

// ---------------------------------------------------------------------------
// Style builders
// ---------------------------------------------------------------------------

func buildElementStyleRequests(objectID string, style *JSONElementStyle) []*slides.Request {
	var reqs []*slides.Request

	// Background color
	if style.BackgroundColor != nil {
		fill := buildSolidFill(style.BackgroundColor)
		if fill != nil {
			reqs = append(reqs, &slides.Request{
				UpdateShapeProperties: &slides.UpdateShapePropertiesRequest{
					ObjectId: objectID,
					ShapeProperties: &slides.ShapeProperties{
						ShapeBackgroundFill: &slides.ShapeBackgroundFill{
							SolidFill: fill,
						},
					},
					Fields: "shapeBackgroundFill.solidFill.color,shapeBackgroundFill.solidFill.alpha",
				},
			})
		}
	}

	// Text style
	if style.TextStyle != nil {
		ts := style.TextStyle
		textStyle := &slides.TextStyle{}
		var fields []string

		if ts.FontSize > 0 {
			textStyle.FontSize = &slides.Dimension{Magnitude: ts.FontSize, Unit: "PT"}
			fields = append(fields, "fontSize")
		}
		if ts.Bold != nil {
			textStyle.Bold = *ts.Bold
			if !*ts.Bold {
				textStyle.ForceSendFields = append(textStyle.ForceSendFields, "Bold")
			}
			fields = append(fields, "bold")
		}
		if ts.Italic != nil {
			textStyle.Italic = *ts.Italic
			if !*ts.Italic {
				textStyle.ForceSendFields = append(textStyle.ForceSendFields, "Italic")
			}
			fields = append(fields, "italic")
		}
		if ts.FontFamily != "" {
			textStyle.FontFamily = ts.FontFamily
			fields = append(fields, "fontFamily")
		}
		if ts.Color != nil {
			textStyle.ForegroundColor = buildOptionalColor(ts.Color)
			fields = append(fields, "foregroundColor")
		}

		if len(fields) > 0 {
			reqs = append(reqs, &slides.Request{
				UpdateTextStyle: &slides.UpdateTextStyleRequest{
					ObjectId: objectID,
					Style:    textStyle,
					TextRange: &slides.Range{
						Type: "ALL",
					},
					Fields: strings.Join(fields, ","),
				},
			})
		}
	}

	return reqs
}

func buildSolidFill(c *JSONColor) *slides.SolidFill {
	alpha := c.Alpha
	if alpha == 0 {
		alpha = 1.0
	}

	if c.Theme != "" {
		return &slides.SolidFill{
			Color: &slides.OpaqueColor{
				ThemeColor: c.Theme,
			},
			Alpha: alpha,
		}
	}

	if len(c.RGB) >= 3 {
		return &slides.SolidFill{
			Color: &slides.OpaqueColor{
				RgbColor: &slides.RgbColor{
					Red:   c.RGB[0] / 255.0,
					Green: c.RGB[1] / 255.0,
					Blue:  c.RGB[2] / 255.0,
				},
			},
			Alpha: alpha,
		}
	}

	return nil
}

func buildOptionalColor(c *JSONColor) *slides.OptionalColor {
	if c.Theme != "" {
		return &slides.OptionalColor{
			OpaqueColor: &slides.OpaqueColor{
				ThemeColor: c.Theme,
			},
		}
	}
	if len(c.RGB) >= 3 {
		return &slides.OptionalColor{
			OpaqueColor: &slides.OpaqueColor{
				RgbColor: &slides.RgbColor{
					Red:   c.RGB[0] / 255.0,
					Green: c.RGB[1] / 255.0,
					Blue:  c.RGB[2] / 255.0,
				},
			},
		}
	}
	return nil
}

// ensureValidID returns the given ID or generates a unique one with the prefix.
func ensureValidID(id, prefix string) string {
	if id != "" {
		return id
	}
	return fmt.Sprintf("%s_%d", prefix, time.Now().UnixNano())
}
