package cmd

import (
	"encoding/json"
	"net/http"
	"strings"
	"testing"
)

func TestExecute_DriveDelete_DefaultAndPermanent(t *testing.T) {
	t.Run("default_trash", func(t *testing.T) {
		origNew := newDriveService
		t.Cleanup(func() { newDriveService = origNew })

		var patchCount int
		svc, closeSrv := newDriveTestService(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
			if !strings.Contains(r.URL.Path, "/files/id1") || (r.Method != http.MethodPatch && r.Method != http.MethodPut) {
				http.NotFound(w, r)
				return
			}
			patchCount++
			requireSupportsAllDrives(t, r)
			body := readBody(t, r)
			if !strings.Contains(body, "\"trashed\":true") {
				t.Fatalf("expected trashed=true body, got: %q", body)
			}
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id":      "id1",
				"trashed": true,
				"kind":    "drive#file",
			})
		}))
		defer closeSrv()

		newDriveService = stubDriveService(svc)

		out := captureStdout(t, func() {
			_ = captureStderr(t, func() {
				if execErr := Execute([]string{"--force", "--account", "a@b.com", "drive", "delete", "id1"}); execErr != nil {
					t.Fatalf("Execute: %v", execErr)
				}
			})
		})
		if !strings.Contains(out, "trashed\ttrue") || !strings.Contains(out, "deleted\tfalse") {
			t.Fatalf("unexpected text output: %q", out)
		}

		jsonOut := captureStdout(t, func() {
			_ = captureStderr(t, func() {
				if execErr := Execute([]string{"--json", "--force", "--account", "a@b.com", "drive", "delete", "id1"}); execErr != nil {
					t.Fatalf("Execute: %v", execErr)
				}
			})
		})
		var parsed struct {
			Trashed bool   `json:"trashed"`
			Deleted bool   `json:"deleted"`
			ID      string `json:"id"`
		}
		if err := json.Unmarshal([]byte(jsonOut), &parsed); err != nil {
			t.Fatalf("json parse: %v\nout=%q", err, jsonOut)
		}
		if !parsed.Trashed || parsed.Deleted || parsed.ID != "id1" {
			t.Fatalf("unexpected json output: %#v", parsed)
		}

		if patchCount != 2 {
			t.Fatalf("expected 2 PATCH calls, got %d", patchCount)
		}
	})

	t.Run("permanent_flag_ignored", func(t *testing.T) {
		origNew := newDriveService
		t.Cleanup(func() { newDriveService = origNew })

		var patchCount int
		svc, closeSrv := newDriveTestService(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
			if !strings.Contains(r.URL.Path, "/files/id1") || (r.Method != http.MethodPatch && r.Method != http.MethodPut) {
				http.NotFound(w, r)
				return
			}
			patchCount++
			requireSupportsAllDrives(t, r)
			body := readBody(t, r)
			if !strings.Contains(body, "\"trashed\":true") {
				t.Fatalf("expected trashed=true body, got: %q", body)
			}
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id":      "id1",
				"trashed": true,
				"kind":    "drive#file",
			})
		}))
		defer closeSrv()

		newDriveService = stubDriveService(svc)

		// --permanent is accepted but ignored; file is trashed, not permanently deleted
		out := captureStdout(t, func() {
			_ = captureStderr(t, func() {
				if execErr := Execute([]string{"--force", "--account", "a@b.com", "drive", "delete", "id1", "--permanent"}); execErr != nil {
					t.Fatalf("Execute: %v", execErr)
				}
			})
		})
		if !strings.Contains(out, "trashed\ttrue") || !strings.Contains(out, "deleted\tfalse") {
			t.Fatalf("unexpected text output: %q", out)
		}

		if patchCount != 1 {
			t.Fatalf("expected 1 PATCH call, got %d", patchCount)
		}
	})
}
