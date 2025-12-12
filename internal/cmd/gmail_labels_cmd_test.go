package cmd

import (
	"context"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"github.com/steipete/gogcli/internal/outfmt"
	"github.com/steipete/gogcli/internal/ui"
	"google.golang.org/api/gmail/v1"
	"google.golang.org/api/option"
)

func TestGmailLabelsGetCmd_JSON(t *testing.T) {
	origNew := newGmailService
	t.Cleanup(func() { newGmailService = origNew })

	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case strings.HasSuffix(r.URL.Path, "/users/me/labels") || strings.HasSuffix(r.URL.Path, "/gmail/v1/users/me/labels"):
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{
				"labels": []map[string]any{
					{"id": "INBOX", "name": "INBOX", "type": "system"},
					{"id": "Label_1", "name": "Custom", "type": "user"},
				},
			})
			return
		case strings.Contains(r.URL.Path, "/users/me/labels/") || strings.Contains(r.URL.Path, "/gmail/v1/users/me/labels/"):
			id := r.URL.Path[strings.LastIndex(r.URL.Path, "/")+1:]
			if id == "inbox" {
				// command should map name->id, but tolerate.
				id = "INBOX"
			}
			if id != "INBOX" {
				http.NotFound(w, r)
				return
			}
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id":             "INBOX",
				"name":           "INBOX",
				"type":           "system",
				"messagesTotal":  123,
				"messagesUnread": 7,
				"threadsTotal":   50,
				"threadsUnread":  3,
			})
			return
		default:
			http.NotFound(w, r)
			return
		}
	}))
	defer srv.Close()

	svc, err := gmail.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(srv.Client()),
		option.WithEndpoint(srv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("NewService: %v", err)
	}
	newGmailService = func(context.Context, string) (*gmail.Service, error) {
		return svc, nil
	}

	flags := &rootFlags{Account: "a@b.com"}

	out := captureStdout(t, func() {
		u, uiErr := ui.New(ui.Options{Stdout: io.Discard, Stderr: io.Discard, Color: "never"})
		if uiErr != nil {
			t.Fatalf("ui.New: %v", uiErr)
		}
		ctx := ui.WithUI(context.Background(), u)
		ctx = outfmt.WithMode(ctx, outfmt.ModeJSON)

		cmd := newGmailLabelsGetCmd(flags)
		cmd.SetContext(ctx)
		cmd.SetArgs([]string{"INBOX"})
		if err := cmd.Execute(); err != nil {
			t.Fatalf("execute: %v", err)
		}
	})

	var parsed struct {
		Label struct {
			ID             string `json:"id"`
			Name           string `json:"name"`
			MessagesTotal  int64  `json:"messagesTotal"`
			MessagesUnread int64  `json:"messagesUnread"`
		} `json:"label"`
	}
	if err := json.Unmarshal([]byte(out), &parsed); err != nil {
		t.Fatalf("json parse: %v\nout=%q", err, out)
	}
	if parsed.Label.ID != "INBOX" || parsed.Label.Name != "INBOX" {
		t.Fatalf("unexpected label: %#v", parsed.Label)
	}
	if parsed.Label.MessagesTotal != 123 || parsed.Label.MessagesUnread != 7 {
		t.Fatalf("unexpected counts: %#v", parsed.Label)
	}
}
