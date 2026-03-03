package cmd

import (
	"context"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/slides/v1"

	"github.com/steipete/gogcli/internal/outfmt"
	"github.com/steipete/gogcli/internal/ui"
)

// jsonPresGetResponse returns a minimal presentation JSON for create-from-json tests.
func jsonPresGetResponse(presID string, slideIDs []string) map[string]any {
	slidesArr := make([]any, 0, len(slideIDs))
	for _, id := range slideIDs {
		slidesArr = append(slidesArr, map[string]any{
			"objectId": id,
			"slideProperties": map[string]any{
				"notesPage": map[string]any{
					"notesProperties": map[string]any{
						"speakerNotesObjectId": "notes_" + id,
					},
					"pageElements": []any{
						map[string]any{
							"objectId": "notes_" + id,
							"shape": map[string]any{
								"placeholder": map[string]any{
									"type": "BODY",
								},
							},
						},
					},
				},
			},
		})
	}

	return map[string]any{
		"presentationId": presID,
		"pageSize": map[string]any{
			"width":  map[string]any{"magnitude": 9144000, "unit": "EMU"},
			"height": map[string]any{"magnitude": 5143500, "unit": "EMU"},
		},
		"slides": slidesArr,
		"layouts": []any{
			map[string]any{
				"objectId": "layout_title_slide",
				"layoutProperties": map[string]any{
					"displayName": "Title slide",
				},
			},
			map[string]any{
				"objectId": "layout_title_and_body",
				"layoutProperties": map[string]any{
					"displayName": "Title and body",
				},
			},
			map[string]any{
				"objectId": "layout_blank",
				"layoutProperties": map[string]any{
					"displayName": "Blank",
				},
			},
		},
	}
}

func TestSlidesCreateFromJSON_Basic(t *testing.T) {
	origSlides := newSlidesService
	origDrive := newDriveService
	t.Cleanup(func() {
		newSlidesService = origSlides
		newDriveService = origDrive
	})

	var batchUpdateCount int
	var capturedSlideIDs []string

	slidesSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch {
		case strings.HasSuffix(r.URL.Path, ":batchUpdate") && r.Method == http.MethodPost:
			batchUpdateCount++
			var req slides.BatchUpdatePresentationRequest
			if err := json.NewDecoder(r.Body).Decode(&req); err == nil {
				for _, rr := range req.Requests {
					if rr.CreateSlide != nil {
						capturedSlideIDs = append(capturedSlideIDs, rr.CreateSlide.ObjectId)
					}
				}
			}
			_ = json.NewEncoder(w).Encode(map[string]any{
				"presentationId": "pres_json_1",
				"replies":        []any{},
			})
		case r.URL.Path == "/v1/presentations" && r.Method == http.MethodPost:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"presentationId": "pres_json_1",
			})
		case strings.Contains(r.URL.Path, "/presentations/pres_json_1") && r.Method == http.MethodGet:
			existingSlides := []string{"default_slide_1"}
			// After first batchUpdate, return the new slides
			if batchUpdateCount > 0 {
				existingSlides = capturedSlideIDs
			}
			_ = json.NewEncoder(w).Encode(jsonPresGetResponse("pres_json_1", existingSlides))
		default:
			http.NotFound(w, r)
		}
	}))
	defer slidesSrv.Close()

	driveSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch {
		case strings.Contains(r.URL.Path, "/files/pres_json_1") && r.Method == http.MethodGet:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id":          "pres_json_1",
				"name":        "Test Deck",
				"webViewLink": "https://docs.google.com/presentation/d/pres_json_1/edit",
			})
		default:
			http.NotFound(w, r)
		}
	}))
	defer driveSrv.Close()

	slidesSvc, err := slides.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(slidesSrv.Client()),
		option.WithEndpoint(slidesSrv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("slides.NewService: %v", err)
	}
	newSlidesService = func(context.Context, string) (*slides.Service, error) { return slidesSvc, nil }

	driveSvc, err := drive.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(driveSrv.Client()),
		option.WithEndpoint(driveSrv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("drive.NewService: %v", err)
	}
	newDriveService = func(context.Context, string) (*drive.Service, error) { return driveSvc, nil }

	// Write JSON spec file
	specJSON := `{
		"presentation": {
			"title": "Test Deck",
			"slides": [
				{
					"layout": "TITLE_SLIDE",
					"content": {
						"CENTERED_TITLE": "Hello World",
						"SUBTITLE": "A test presentation"
					}
				},
				{
					"layout": "TITLE_AND_BODY",
					"content": {
						"TITLE": "Slide Two",
						"BODY": "Some body text with **bold**"
					}
				}
			]
		}
	}`
	specPath := filepath.Join(t.TempDir(), "spec.json")
	if err := os.WriteFile(specPath, []byte(specJSON), 0o644); err != nil {
		t.Fatalf("write spec: %v", err)
	}

	flags := &RootFlags{Account: "a@b.com"}

	out := captureStdout(t, func() {
		u, uiErr := ui.New(ui.Options{Stdout: os.Stdout, Stderr: io.Discard, Color: "never"})
		if uiErr != nil {
			t.Fatalf("ui.New: %v", uiErr)
		}
		ctx := ui.WithUI(context.Background(), u)

		cmd := &SlidesCreateFromJSONCmd{
			Title:       "Test Deck",
			ContentFile: specPath,
		}
		if err := cmd.Run(ctx, flags); err != nil {
			t.Fatalf("Run: %v", err)
		}
	})

	if !strings.Contains(out, "Created presentation with 2 slides") {
		t.Errorf("expected slide count message, got: %q", out)
	}
	if !strings.Contains(out, "id\tpres_json_1") {
		t.Errorf("expected presentation ID, got: %q", out)
	}
	if !strings.Contains(out, "link\thttps://docs.google.com/presentation/d/pres_json_1/edit") {
		t.Errorf("expected presentation link, got: %q", out)
	}
	if len(capturedSlideIDs) != 2 {
		t.Errorf("expected 2 slides created, got %d", len(capturedSlideIDs))
	}
	// No notes → only 1 batchUpdate
	if batchUpdateCount != 1 {
		t.Errorf("expected 1 batchUpdate (no notes), got %d", batchUpdateCount)
	}
}

func TestSlidesCreateFromJSON_WithNotes(t *testing.T) {
	origSlides := newSlidesService
	origDrive := newDriveService
	t.Cleanup(func() {
		newSlidesService = origSlides
		newDriveService = origDrive
	})

	var batchUpdateCount int
	var capturedSlideIDs []string

	slidesSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch {
		case strings.HasSuffix(r.URL.Path, ":batchUpdate") && r.Method == http.MethodPost:
			batchUpdateCount++
			var req slides.BatchUpdatePresentationRequest
			if err := json.NewDecoder(r.Body).Decode(&req); err == nil {
				for _, rr := range req.Requests {
					if rr.CreateSlide != nil {
						capturedSlideIDs = append(capturedSlideIDs, rr.CreateSlide.ObjectId)
					}
				}
			}
			_ = json.NewEncoder(w).Encode(map[string]any{
				"presentationId": "pres_notes",
				"replies":        []any{},
			})
		case r.URL.Path == "/v1/presentations" && r.Method == http.MethodPost:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"presentationId": "pres_notes",
			})
		case strings.Contains(r.URL.Path, "/presentations/pres_notes") && r.Method == http.MethodGet:
			existingSlides := []string{"default_slide_1"}
			if batchUpdateCount > 0 {
				existingSlides = capturedSlideIDs
			}
			_ = json.NewEncoder(w).Encode(jsonPresGetResponse("pres_notes", existingSlides))
		default:
			http.NotFound(w, r)
		}
	}))
	defer slidesSrv.Close()

	driveSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch {
		case strings.Contains(r.URL.Path, "/files/pres_notes") && r.Method == http.MethodGet:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id":          "pres_notes",
				"name":        "Notes Deck",
				"webViewLink": "https://docs.google.com/presentation/d/pres_notes/edit",
			})
		default:
			http.NotFound(w, r)
		}
	}))
	defer driveSrv.Close()

	slidesSvc, err := slides.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(slidesSrv.Client()),
		option.WithEndpoint(slidesSrv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("slides.NewService: %v", err)
	}
	newSlidesService = func(context.Context, string) (*slides.Service, error) { return slidesSvc, nil }

	driveSvc, err := drive.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(driveSrv.Client()),
		option.WithEndpoint(driveSrv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("drive.NewService: %v", err)
	}
	newDriveService = func(context.Context, string) (*drive.Service, error) { return driveSvc, nil }

	flags := &RootFlags{Account: "a@b.com"}
	u, uiErr := ui.New(ui.Options{Stdout: io.Discard, Stderr: io.Discard, Color: "never"})
	if uiErr != nil {
		t.Fatalf("ui.New: %v", uiErr)
	}
	ctx := ui.WithUI(context.Background(), u)

	cmd := &SlidesCreateFromJSONCmd{
		Title: "Notes Deck",
		Content: `{
			"presentation": {
				"title": "Notes Deck",
				"slides": [
					{
						"layout": "TITLE_AND_BODY",
						"content": {"TITLE": "Slide 1"},
						"notes": "Speaker notes for slide 1"
					}
				]
			}
		}`,
	}

	_ = captureStdout(t, func() {
		if err := cmd.Run(ctx, flags); err != nil {
			t.Fatalf("Run: %v", err)
		}
	})

	// With notes: 1 for creating slides + 1 for inserting notes
	if batchUpdateCount != 2 {
		t.Errorf("expected 2 batchUpdates (with notes), got %d", batchUpdateCount)
	}
}

func TestSlidesCreateFromJSON_JSONOutput(t *testing.T) {
	origSlides := newSlidesService
	origDrive := newDriveService
	t.Cleanup(func() {
		newSlidesService = origSlides
		newDriveService = origDrive
	})

	slidesSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch {
		case strings.HasSuffix(r.URL.Path, ":batchUpdate") && r.Method == http.MethodPost:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"presentationId": "pres_jout",
				"replies":        []any{},
			})
		case r.URL.Path == "/v1/presentations" && r.Method == http.MethodPost:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"presentationId": "pres_jout",
			})
		case strings.Contains(r.URL.Path, "/presentations/pres_jout") && r.Method == http.MethodGet:
			_ = json.NewEncoder(w).Encode(jsonPresGetResponse("pres_jout", []string{"default_1"}))
		default:
			http.NotFound(w, r)
		}
	}))
	defer slidesSrv.Close()

	driveSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch {
		case strings.Contains(r.URL.Path, "/files/pres_jout") && r.Method == http.MethodGet:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id":          "pres_jout",
				"name":        "JSON Out",
				"webViewLink": "https://docs.google.com/presentation/d/pres_jout/edit",
			})
		default:
			http.NotFound(w, r)
		}
	}))
	defer driveSrv.Close()

	slidesSvc, err := slides.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(slidesSrv.Client()),
		option.WithEndpoint(slidesSrv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("slides.NewService: %v", err)
	}
	newSlidesService = func(context.Context, string) (*slides.Service, error) { return slidesSvc, nil }

	driveSvc, err := drive.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(driveSrv.Client()),
		option.WithEndpoint(driveSrv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("drive.NewService: %v", err)
	}
	newDriveService = func(context.Context, string) (*drive.Service, error) { return driveSvc, nil }

	flags := &RootFlags{Account: "a@b.com"}
	u, uiErr := ui.New(ui.Options{Stdout: io.Discard, Stderr: io.Discard, Color: "never"})
	if uiErr != nil {
		t.Fatalf("ui.New: %v", uiErr)
	}
	ctx := ui.WithUI(context.Background(), u)
	ctx = outfmt.WithMode(ctx, outfmt.Mode{JSON: true})

	out := captureStdout(t, func() {
		cmd := &SlidesCreateFromJSONCmd{
			Title: "JSON Out",
			Content: `{
				"presentation": {
					"title": "JSON Out",
					"slides": [
						{"layout": "BLANK"}
					]
				}
			}`,
		}
		if err := cmd.Run(ctx, flags); err != nil {
			t.Fatalf("Run: %v", err)
		}
	})

	var result map[string]any
	if err := json.Unmarshal([]byte(out), &result); err != nil {
		t.Fatalf("JSON parse: %v\noutput: %q", err, out)
	}
	if result["presentationId"] != "pres_jout" {
		t.Errorf("expected presentationId=pres_jout, got %v", result["presentationId"])
	}
	file, ok := result["file"].(map[string]any)
	if !ok {
		t.Fatalf("expected file object in output")
	}
	if file["webViewLink"] != "https://docs.google.com/presentation/d/pres_jout/edit" {
		t.Errorf("unexpected webViewLink: %v", file["webViewLink"])
	}
}

func TestSlidesCreateFromJSON_EmptyTitle(t *testing.T) {
	cmd := &SlidesCreateFromJSONCmd{
		Content: `{"presentation":{"title":"","slides":[{"layout":"BLANK"}]}}`,
	}
	flags := &RootFlags{Account: "a@b.com"}
	u, _ := ui.New(ui.Options{Stdout: io.Discard, Stderr: io.Discard, Color: "never"})
	ctx := ui.WithUI(context.Background(), u)

	err := cmd.Run(ctx, flags)
	if err == nil || !strings.Contains(err.Error(), "empty title") {
		t.Fatalf("expected empty title error, got: %v", err)
	}
}

func TestSlidesCreateFromJSON_NoContent(t *testing.T) {
	cmd := &SlidesCreateFromJSONCmd{}
	flags := &RootFlags{Account: "a@b.com"}
	u, _ := ui.New(ui.Options{Stdout: io.Discard, Stderr: io.Discard, Color: "never"})
	ctx := ui.WithUI(context.Background(), u)

	err := cmd.Run(ctx, flags)
	if err == nil || !strings.Contains(err.Error(), "either --content or --content-file") {
		t.Fatalf("expected content required error, got: %v", err)
	}
}

func TestSlidesCreateFromJSON_NoSlides(t *testing.T) {
	cmd := &SlidesCreateFromJSONCmd{
		Title:   "Empty",
		Content: `{"presentation":{"title":"Empty","slides":[]}}`,
	}
	flags := &RootFlags{Account: "a@b.com"}
	u, _ := ui.New(ui.Options{Stdout: io.Discard, Stderr: io.Discard, Color: "never"})
	ctx := ui.WithUI(context.Background(), u)

	err := cmd.Run(ctx, flags)
	if err == nil || !strings.Contains(err.Error(), "at least one slide") {
		t.Fatalf("expected at least one slide error, got: %v", err)
	}
}

func TestSlidesCreateFromJSON_ContentFile(t *testing.T) {
	origSlides := newSlidesService
	origDrive := newDriveService
	t.Cleanup(func() {
		newSlidesService = origSlides
		newDriveService = origDrive
	})

	slidesSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch {
		case strings.HasSuffix(r.URL.Path, ":batchUpdate") && r.Method == http.MethodPost:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"presentationId": "pres_cf",
				"replies":        []any{},
			})
		case r.URL.Path == "/v1/presentations" && r.Method == http.MethodPost:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"presentationId": "pres_cf",
			})
		case strings.Contains(r.URL.Path, "/presentations/pres_cf") && r.Method == http.MethodGet:
			_ = json.NewEncoder(w).Encode(jsonPresGetResponse("pres_cf", []string{"default_1"}))
		default:
			http.NotFound(w, r)
		}
	}))
	defer slidesSrv.Close()

	driveSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")

		switch {
		case strings.Contains(r.URL.Path, "/files/pres_cf") && r.Method == http.MethodGet:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id":          "pres_cf",
				"name":        "From File",
				"webViewLink": "https://docs.google.com/presentation/d/pres_cf/edit",
			})
		default:
			http.NotFound(w, r)
		}
	}))
	defer driveSrv.Close()

	slidesSvc, err := slides.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(slidesSrv.Client()),
		option.WithEndpoint(slidesSrv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("slides.NewService: %v", err)
	}
	newSlidesService = func(context.Context, string) (*slides.Service, error) { return slidesSvc, nil }

	driveSvc, err := drive.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(driveSrv.Client()),
		option.WithEndpoint(driveSrv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("drive.NewService: %v", err)
	}
	newDriveService = func(context.Context, string) (*drive.Service, error) { return driveSvc, nil }

	specPath := filepath.Join(t.TempDir(), "deck.json")
	specJSON := `{
		"presentation": {
			"title": "From File",
			"slides": [{"layout": "TITLE_ONLY", "content": {"TITLE": "Hello"}}]
		}
	}`
	if err := os.WriteFile(specPath, []byte(specJSON), 0o644); err != nil {
		t.Fatalf("write spec: %v", err)
	}

	flags := &RootFlags{Account: "a@b.com"}

	out := captureStdout(t, func() {
		u, uiErr := ui.New(ui.Options{Stdout: os.Stdout, Stderr: io.Discard, Color: "never"})
		if uiErr != nil {
			t.Fatalf("ui.New: %v", uiErr)
		}
		ctx := ui.WithUI(context.Background(), u)

		cmd := &SlidesCreateFromJSONCmd{
			ContentFile: specPath,
		}
		if err := cmd.Run(ctx, flags); err != nil {
			t.Fatalf("Run: %v", err)
		}
	})

	// Title comes from JSON spec
	if !strings.Contains(out, "id\tpres_cf") {
		t.Errorf("expected presentation ID, got: %q", out)
	}
}

func TestParseMarkdownText(t *testing.T) {
	tests := []struct {
		name       string
		input      string
		wantPlain  string
		wantStyles []string
	}{
		{
			name:       "plain text",
			input:      "Hello world",
			wantPlain:  "Hello world",
			wantStyles: nil,
		},
		{
			name:       "bold text",
			input:      "Hello **world**",
			wantPlain:  "Hello world",
			wantStyles: []string{"bold"},
		},
		{
			name:       "italic text",
			input:      "Hello *world*",
			wantPlain:  "Hello world",
			wantStyles: []string{"italic"},
		},
		{
			name:       "strikethrough",
			input:      "Hello ~~world~~",
			wantPlain:  "Hello world",
			wantStyles: []string{"strikethrough"},
		},
		{
			name:       "link",
			input:      "Visit [Google](https://google.com)",
			wantPlain:  "Visit Google",
			wantStyles: []string{"link"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			plain, annotations := parseMarkdownText(tt.input)
			if plain != tt.wantPlain {
				t.Errorf("plain text = %q, want %q", plain, tt.wantPlain)
			}
			if len(annotations) != len(tt.wantStyles) {
				t.Errorf("got %d annotations, want %d", len(annotations), len(tt.wantStyles))
				return
			}
			for i, a := range annotations {
				if a.style != tt.wantStyles[i] {
					t.Errorf("annotation[%d].style = %q, want %q", i, a.style, tt.wantStyles[i])
				}
			}
		})
	}
}

func TestParsePlaceholderType(t *testing.T) {
	tests := []struct {
		key       string
		wantType  string
		wantIndex int
	}{
		{"TITLE", "TITLE", -1},
		{"BODY", "BODY", -1},
		{"BODY_0", "BODY", 0},
		{"BODY_1", "BODY", 1},
		{"CENTERED_TITLE", "CENTERED_TITLE", -1},
		{"SUBTITLE", "SUBTITLE", -1},
	}

	for _, tt := range tests {
		t.Run(tt.key, func(t *testing.T) {
			typ, idx := parsePlaceholderType(tt.key)
			if typ != tt.wantType {
				t.Errorf("type = %q, want %q", typ, tt.wantType)
			}
			if idx != tt.wantIndex {
				t.Errorf("index = %d, want %d", idx, tt.wantIndex)
			}
		})
	}
}

func TestJSONContentValueUnmarshal(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		wantType string
		wantText string
	}{
		{
			name:     "plain string",
			input:    `"Hello World"`,
			wantType: "text",
			wantText: "Hello World",
		},
		{
			name:     "text object",
			input:    `{"type": "text", "text": "Formatted **text**"}`,
			wantType: "text",
			wantText: "Formatted **text**",
		},
		{
			name:     "table object",
			input:    `{"type": "table", "rows": 2, "columns": 3, "data": [["a","b","c"],["d","e","f"]]}`,
			wantType: "table",
		},
		{
			name:     "image object",
			input:    `{"type": "image", "url": "https://example.com/img.png"}`,
			wantType: "image",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var v JSONContentValue
			if err := json.Unmarshal([]byte(tt.input), &v); err != nil {
				t.Fatalf("unmarshal: %v", err)
			}
			if v.Type != tt.wantType {
				t.Errorf("type = %q, want %q", v.Type, tt.wantType)
			}
			if tt.wantText != "" && v.Text != tt.wantText {
				t.Errorf("text = %q, want %q", v.Text, tt.wantText)
			}
		})
	}
}

func TestSlidesCreateFromJSON_DryRun(t *testing.T) {
	origSlides := newSlidesService
	origDrive := newDriveService
	t.Cleanup(func() {
		newSlidesService = origSlides
		newDriveService = origDrive
	})

	// Services should NOT be called in dry-run mode
	newSlidesService = func(context.Context, string) (*slides.Service, error) {
		t.Fatal("slides service should not be created in dry-run mode")
		return nil, context.Canceled
	}
	newDriveService = func(context.Context, string) (*drive.Service, error) {
		t.Fatal("drive service should not be created in dry-run mode")
		return nil, context.Canceled
	}

	flags := &RootFlags{Account: "a@b.com", DryRun: true}

	out := captureStdout(t, func() {
		u, uiErr := ui.New(ui.Options{Stdout: os.Stdout, Stderr: io.Discard, Color: "never"})
		if uiErr != nil {
			t.Fatalf("ui.New: %v", uiErr)
		}
		ctx := ui.WithUI(context.Background(), u)

		cmd := &SlidesCreateFromJSONCmd{
			Title: "Dry Run Test",
			Content: `{
				"presentation": {
					"title": "Dry Run Test",
					"slides": [{"layout": "BLANK"}]
				}
			}`,
		}
		err := cmd.Run(ctx, flags)
		// dryRunExit returns ExitError{Code: 0}
		if err != nil {
			if exitErr, ok := err.(*ExitError); ok && exitErr.Code == 0 {
				// expected
			} else {
				t.Fatalf("unexpected error: %v", err)
			}
		}
	})

	if !strings.Contains(out, "Dry run") {
		t.Errorf("expected dry-run output, got: %q", out)
	}
}

func TestSlidesCreateFromJSON_DryRunJSON(t *testing.T) {
	origSlides := newSlidesService
	origDrive := newDriveService
	t.Cleanup(func() {
		newSlidesService = origSlides
		newDriveService = origDrive
	})

	newSlidesService = func(context.Context, string) (*slides.Service, error) {
		t.Fatal("slides service should not be created in dry-run mode")
		return nil, context.Canceled
	}
	newDriveService = func(context.Context, string) (*drive.Service, error) {
		t.Fatal("drive service should not be created in dry-run mode")
		return nil, context.Canceled
	}

	flags := &RootFlags{Account: "a@b.com", DryRun: true}
	u, _ := ui.New(ui.Options{Stdout: io.Discard, Stderr: io.Discard, Color: "never"})
	ctx := ui.WithUI(context.Background(), u)
	ctx = outfmt.WithMode(ctx, outfmt.Mode{JSON: true})

	out := captureStdout(t, func() {
		cmd := &SlidesCreateFromJSONCmd{
			Title: "Dry JSON",
			Content: `{
				"presentation": {
					"title": "Dry JSON",
					"slides": [{"layout": "TITLE_AND_BODY", "content": {"TITLE": "Hi"}}]
				}
			}`,
		}
		err := cmd.Run(ctx, flags)
		if err != nil {
			if exitErr, ok := err.(*ExitError); ok && exitErr.Code == 0 {
				// expected
			} else {
				t.Fatalf("unexpected error: %v", err)
			}
		}
	})

	var result map[string]any
	if err := json.Unmarshal([]byte(out), &result); err != nil {
		t.Fatalf("JSON parse: %v\noutput: %q", err, out)
	}
	if result["dry_run"] != true {
		t.Errorf("expected dry_run=true, got %v", result["dry_run"])
	}
	if result["op"] != "slides.create-from-json" {
		t.Errorf("expected op=slides.create-from-json, got %v", result["op"])
	}
	req, ok := result["request"].(map[string]any)
	if !ok {
		t.Fatal("expected request object")
	}
	if req["title"] != "Dry JSON" {
		t.Errorf("expected title=Dry JSON, got %v", req["title"])
	}
	if req["slides"] != float64(1) {
		t.Errorf("expected slides=1, got %v", req["slides"])
	}
}

func TestNormalizePredefinedLayout(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"TITLE_SLIDE", "TITLE"},
		{"CAPTION", "CAPTION_ONLY"},
		{"TITLE_AND_BODY", "TITLE_AND_BODY"},
		{"SECTION_HEADER", "SECTION_HEADER"},
		{"BLANK", "BLANK"},
		{"TITLE_ONLY", "TITLE_ONLY"},
		{"BIG_NUMBER", "BIG_NUMBER"},
		{"MAIN_POINT", "MAIN_POINT"},
		{"ONE_COLUMN_TEXT", "ONE_COLUMN_TEXT"},
		{"TITLE_AND_TWO_COLUMNS", "TITLE_AND_TWO_COLUMNS"},
		{"SECTION_TITLE_AND_DESCRIPTION", "SECTION_TITLE_AND_DESCRIPTION"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := normalizePredefinedLayout(tt.input)
			if got != tt.want {
				t.Errorf("normalizePredefinedLayout(%q) = %q, want %q", tt.input, got, tt.want)
			}
		})
	}
}
