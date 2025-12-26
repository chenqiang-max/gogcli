package cmd

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/steipete/gogcli/internal/config"
)

func TestExecute_AuthCredentials_JSON(t *testing.T) {
	home := t.TempDir()
	t.Setenv("HOME", home)

	in := filepath.Join(t.TempDir(), "creds.json")
	if err := os.WriteFile(in, []byte(`{"installed":{"client_id":"id","client_secret":"sec"}}`), 0o600); err != nil {
		t.Fatalf("WriteFile: %v", err)
	}

	out := captureStdout(t, func() {
		_ = captureStderr(t, func() {
			if err := Execute([]string{"--json", "auth", "credentials", in}); err != nil {
				t.Fatalf("Execute: %v", err)
			}
		})
	})

	var parsed struct {
		Saved bool   `json:"saved"`
		Path  string `json:"path"`
	}
	if err := json.Unmarshal([]byte(out), &parsed); err != nil {
		t.Fatalf("json parse: %v\nout=%q", err, out)
	}
	if !parsed.Saved || parsed.Path == "" {
		t.Fatalf("unexpected: %#v", parsed)
	}
	outPath, err := config.ClientCredentialsPath()
	if err != nil {
		t.Fatalf("ClientCredentialsPath: %v", err)
	}
	if parsed.Path != outPath {
		t.Fatalf("expected %q, got %q", outPath, parsed.Path)
	}
	if st, err := os.Stat(outPath); err != nil || st.Size() == 0 {
		t.Fatalf("stat: %v size=%d", err, st.Size())
	}
}

func TestExecute_AuthCredentials_Stdin_JSON(t *testing.T) {
	home := t.TempDir()
	t.Setenv("HOME", home)

	out := captureStdout(t, func() {
		_ = captureStderr(t, func() {
			withStdin(t, `{"installed":{"client_id":"id","client_secret":"sec"}}`, func() {
				if err := Execute([]string{"--json", "auth", "credentials", "-"}); err != nil {
					t.Fatalf("Execute: %v", err)
				}
			})
		})
	})

	var parsed struct {
		Saved bool   `json:"saved"`
		Path  string `json:"path"`
	}
	if err := json.Unmarshal([]byte(out), &parsed); err != nil {
		t.Fatalf("json parse: %v\nout=%q", err, out)
	}
	if !parsed.Saved || parsed.Path == "" {
		t.Fatalf("unexpected: %#v", parsed)
	}
}
