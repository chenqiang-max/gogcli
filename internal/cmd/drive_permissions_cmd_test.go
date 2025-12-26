package cmd

import (
	"bytes"
	"context"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"github.com/steipete/gogcli/internal/outfmt"
	"github.com/steipete/gogcli/internal/ui"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
)

func TestDrivePermissionsCmd_TextAndJSON(t *testing.T) {
	origNew := newDriveService
	t.Cleanup(func() { newDriveService = origNew })

	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == http.MethodGet && strings.Contains(r.URL.Path, "/files/id1/permissions"):
			if r.URL.Query().Get("pageSize") != "1" {
				t.Fatalf("expected pageSize=1, got: %q", r.URL.RawQuery)
			}
			if r.URL.Query().Get("pageToken") != "p1" {
				t.Fatalf("expected pageToken=p1, got: %q", r.URL.RawQuery)
			}
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{
				"permissions": []map[string]any{
					{"id": "p1", "type": "anyone", "role": "reader", "emailAddress": "a@b.com"},
				},
				"nextPageToken": "npt",
			})
			return
		default:
			http.NotFound(w, r)
			return
		}
	}))
	defer srv.Close()

	svc, err := drive.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(srv.Client()),
		option.WithEndpoint(srv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("NewService: %v", err)
	}
	newDriveService = func(context.Context, string) (*drive.Service, error) { return svc, nil }

	flags := &rootFlags{Account: "a@b.com"}

	// Text mode: table to stdout + next page hint to stderr.
	var errBuf bytes.Buffer
	u, err := ui.New(ui.Options{Stdout: io.Discard, Stderr: &errBuf, Color: "never"})
	if err != nil {
		t.Fatalf("ui.New: %v", err)
	}
	ctx := ui.WithUI(context.Background(), u)
	ctx = outfmt.WithMode(ctx, outfmt.Mode{})

	textOut := captureStdout(t, func() {
		cmd := newDrivePermissionsCmd(flags)
		cmd.SetContext(ctx)
		cmd.SetArgs([]string{"--max", "1", "--page", "p1", "id1"})
		if execErr := cmd.Execute(); execErr != nil {
			t.Fatalf("execute: %v", execErr)
		}
	})
	if !strings.Contains(textOut, "ID") || !strings.Contains(textOut, "TYPE") {
		t.Fatalf("unexpected table header: %q", textOut)
	}
	if !strings.Contains(textOut, "p1") || !strings.Contains(textOut, "anyone") || !strings.Contains(textOut, "reader") {
		t.Fatalf("missing permission row: %q", textOut)
	}
	if !strings.Contains(errBuf.String(), "--page npt") {
		t.Fatalf("missing next page hint: %q", errBuf.String())
	}

	// JSON mode: JSON to stdout and no next-page hint to stderr.
	var errBuf2 bytes.Buffer
	u2, err := ui.New(ui.Options{Stdout: io.Discard, Stderr: &errBuf2, Color: "never"})
	if err != nil {
		t.Fatalf("ui.New: %v", err)
	}
	ctx2 := ui.WithUI(context.Background(), u2)
	ctx2 = outfmt.WithMode(ctx2, outfmt.Mode{JSON: true})

	jsonOut := captureStdout(t, func() {
		cmd := newDrivePermissionsCmd(flags)
		cmd.SetContext(ctx2)
		cmd.SetArgs([]string{"--max", "1", "--page", "p1", "id1"})
		if execErr := cmd.Execute(); execErr != nil {
			t.Fatalf("execute: %v", execErr)
		}
	})
	if errBuf2.String() != "" {
		t.Fatalf("expected no stderr in json mode, got: %q", errBuf2.String())
	}

	var parsed struct {
		FileID          string              `json:"fileId"`
		PermissionCount int                 `json:"permissionCount"`
		Permissions     []*drive.Permission `json:"permissions"`
		NextPageToken   string              `json:"nextPageToken"`
	}
	if err := json.Unmarshal([]byte(jsonOut), &parsed); err != nil {
		t.Fatalf("json parse: %v\nout=%q", err, jsonOut)
	}
	if parsed.FileID != "id1" || parsed.NextPageToken != "npt" || parsed.PermissionCount != 1 || len(parsed.Permissions) != 1 {
		t.Fatalf("unexpected json: %#v", parsed)
	}
}
