package cmd

import (
	"context"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"google.golang.org/api/calendar/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/people/v1"
)

func TestExecute_ContactsList_Text(t *testing.T) {
	origNew := newPeopleContactsService
	t.Cleanup(func() { newPeopleContactsService = origNew })

	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if !strings.Contains(r.URL.Path, "/people/me/connections") {
			http.NotFound(w, r)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]any{
			"connections": []map[string]any{
				{
					"resourceName": "people/c1",
					"names":        []map[string]any{{"displayName": "Ada"}},
					"emailAddresses": []map[string]any{
						{"value": "ada@example.com"},
					},
				},
			},
			"nextPageToken": "npt",
		})
	}))
	defer srv.Close()

	svc, err := people.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(srv.Client()),
		option.WithEndpoint(srv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("NewService: %v", err)
	}
	newPeopleContactsService = func(context.Context, string) (*people.Service, error) { return svc, nil }

	out := captureStdout(t, func() {
		errOut := captureStderr(t, func() {
			if err := Execute([]string{"--account", "a@b.com", "contacts", "list", "--max", "1"}); err != nil {
				t.Fatalf("Execute: %v", err)
			}
		})
		if !strings.Contains(errOut, "# Next page: --page npt") {
			t.Fatalf("unexpected stderr=%q", errOut)
		}
	})
	if !strings.Contains(out, "RESOURCE") || !strings.Contains(out, "people/c1") || !strings.Contains(out, "Ada") {
		t.Fatalf("unexpected out=%q", out)
	}
}

func TestExecute_ContactsGet_ByResource_Text(t *testing.T) {
	origNew := newPeopleContactsService
	t.Cleanup(func() { newPeopleContactsService = origNew })

	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if !(strings.Contains(r.URL.Path, "/people/c1") && r.Method == http.MethodGet) {
			http.NotFound(w, r)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]any{
			"resourceName": "people/c1",
			"names":        []map[string]any{{"displayName": "Ada"}},
			"emailAddresses": []map[string]any{
				{"value": "ada@example.com"},
			},
			"phoneNumbers": []map[string]any{{"value": "+1"}},
		})
	}))
	defer srv.Close()

	svc, err := people.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(srv.Client()),
		option.WithEndpoint(srv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("NewService: %v", err)
	}
	newPeopleContactsService = func(context.Context, string) (*people.Service, error) { return svc, nil }

	out := captureStdout(t, func() {
		_ = captureStderr(t, func() {
			if err := Execute([]string{"--account", "a@b.com", "contacts", "get", "people/c1"}); err != nil {
				t.Fatalf("Execute: %v", err)
			}
		})
	})
	if !strings.Contains(out, "resource\tpeople/c1") || !strings.Contains(out, "email\tada@example.com") {
		t.Fatalf("unexpected out=%q", out)
	}
}

func TestExecute_CalendarFreeBusy_Text(t *testing.T) {
	origNew := newCalendarService
	t.Cleanup(func() { newCalendarService = origNew })

	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if !(strings.Contains(r.URL.Path, "/freeBusy") && r.Method == http.MethodPost) {
			http.NotFound(w, r)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]any{
			"calendars": map[string]any{
				"c1": map[string]any{
					"busy": []map[string]any{{"start": "2025-12-17T10:00:00Z", "end": "2025-12-17T11:00:00Z"}},
				},
			},
		})
	}))
	defer srv.Close()

	svc, err := calendar.NewService(context.Background(),
		option.WithoutAuthentication(),
		option.WithHTTPClient(srv.Client()),
		option.WithEndpoint(srv.URL+"/"),
	)
	if err != nil {
		t.Fatalf("NewService: %v", err)
	}
	newCalendarService = func(context.Context, string) (*calendar.Service, error) { return svc, nil }

	out := captureStdout(t, func() {
		_ = captureStderr(t, func() {
			if err := Execute([]string{"--account", "a@b.com", "calendar", "freebusy", "c1", "--from", "2025-12-17T00:00:00Z", "--to", "2025-12-18T00:00:00Z"}); err != nil {
				t.Fatalf("Execute: %v", err)
			}
		})
	})
	if !strings.Contains(out, "CALENDAR") || !strings.Contains(out, "c1") || !strings.Contains(out, "2025-12-17T10:00:00Z") {
		t.Fatalf("unexpected out=%q", out)
	}
}
