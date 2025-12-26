package outfmt

import (
	"bytes"
	"context"
	"testing"
)

func TestFromFlags(t *testing.T) {
	if _, err := FromFlags(true, true); err == nil {
		t.Fatalf("expected error when combining --json and --plain")
	}
	got, err := FromFlags(true, false)
	if err != nil {
		t.Fatalf("unexpected err: %v", err)
	}
	if !got.JSON || got.Plain {
		t.Fatalf("unexpected mode: %#v", got)
	}
}

func TestContextMode(t *testing.T) {
	ctx := context.Background()
	if IsJSON(ctx) || IsPlain(ctx) {
		t.Fatalf("expected default text")
	}
	ctx = WithMode(ctx, Mode{JSON: true})
	if !IsJSON(ctx) || IsPlain(ctx) {
		t.Fatalf("expected json-only")
	}
}

func TestWriteJSON(t *testing.T) {
	var buf bytes.Buffer
	if err := WriteJSON(&buf, map[string]any{"ok": true}); err != nil {
		t.Fatalf("err: %v", err)
	}
	if buf.Len() == 0 {
		t.Fatalf("expected output")
	}
}
