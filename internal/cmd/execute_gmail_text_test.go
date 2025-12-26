package cmd

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"google.golang.org/api/gmail/v1"
	"google.golang.org/api/option"
)

func TestExecute_GmailThread_Text_Download(t *testing.T) {
	origNew := newGmailService
	t.Cleanup(func() { newGmailService = origNew })

	t.Setenv("XDG_CONFIG_HOME", t.TempDir())

	attData := []byte("hello")
	attEncoded := base64.RawURLEncoding.EncodeToString(attData)
	bodyEncoded := base64.RawURLEncoding.EncodeToString([]byte("body"))

	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case strings.Contains(r.URL.Path, "/gmail/v1/users/me/threads/t-thread-1"):
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id": "t-thread-1",
				"messages": []map[string]any{
					{
						"id": "m-thread-1",
						"payload": map[string]any{
							"headers": []map[string]any{
								{"name": "From", "value": "Me <me@example.com>"},
								{"name": "To", "value": "You <you@example.com>"},
								{"name": "Subject", "value": "Hello"},
								{"name": "Date", "value": "Wed, 17 Dec 2025 14:00:00 -0800"},
							},
							"parts": []map[string]any{
								{ // body
									"mimeType": "text/plain",
									"body":     map[string]any{"data": bodyEncoded},
								},
								{ // attachment
									"filename": "a.txt",
									"mimeType": "text/plain",
									"body":     map[string]any{"attachmentId": "a-thread-1", "size": len(attData)},
								},
							},
						},
					},
				},
			})
			return
		case strings.Contains(r.URL.Path, "/gmail/v1/users/me/messages/m-thread-1/attachments/a-thread-1"):
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{"data": attEncoded})
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
	newGmailService = func(context.Context, string) (*gmail.Service, error) { return svc, nil }

	out := captureStdout(t, func() {
		_ = captureStderr(t, func() {
			if err := Execute([]string{"--account", "a@b.com", "gmail", "thread", "t-thread-1", "--download"}); err != nil {
				t.Fatalf("Execute: %v", err)
			}
		})
	})
	if !strings.Contains(out, "Message: m-thread-1") || !strings.Contains(out, "Attachments:") || !(strings.Contains(out, "Saved:") || strings.Contains(out, "Cached:")) {
		t.Fatalf("unexpected out=%q", out)
	}
}

func TestExecute_GmailDraftsGet_Text_Download(t *testing.T) {
	origNew := newGmailService
	t.Cleanup(func() { newGmailService = origNew })

	t.Setenv("XDG_CONFIG_HOME", t.TempDir())

	attData := []byte("hello")
	attEncoded := base64.RawURLEncoding.EncodeToString(attData)

	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case strings.Contains(r.URL.Path, "/gmail/v1/users/me/drafts/d1"):
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id": "d1",
				"message": map[string]any{
					"id": "m-draft-1",
					"payload": map[string]any{
						"headers": []map[string]any{
							{"name": "To", "value": "x@y.com"},
							{"name": "Subject", "value": "S"},
						},
						"parts": []map[string]any{
							{
								"filename": "a.txt",
								"mimeType": "text/plain",
								"body":     map[string]any{"attachmentId": "a-draft-1", "size": len(attData)},
							},
						},
					},
				},
			})
			return
		case strings.Contains(r.URL.Path, "/gmail/v1/users/me/messages/m-draft-1/attachments/a-draft-1"):
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{"data": attEncoded})
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
	newGmailService = func(context.Context, string) (*gmail.Service, error) { return svc, nil }

	out := captureStdout(t, func() {
		_ = captureStderr(t, func() {
			if err := Execute([]string{"--account", "a@b.com", "gmail", "drafts", "get", "d1", "--download"}); err != nil {
				t.Fatalf("Execute: %v", err)
			}
		})
	})
	if !strings.Contains(out, "Draft-ID: d1") || !strings.Contains(out, "Attachments:") || (!strings.Contains(out, "Saved:") && !strings.Contains(out, "Cached:")) {
		t.Fatalf("unexpected out=%q", out)
	}
}
