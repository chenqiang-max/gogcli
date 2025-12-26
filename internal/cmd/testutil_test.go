package cmd

import (
	"io"
	"os"
	"testing"
)

func captureStdout(t *testing.T, fn func()) string {
	t.Helper()

	orig := os.Stdout
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatalf("pipe: %v", err)
	}
	os.Stdout = w

	fn()

	_ = w.Close()
	os.Stdout = orig
	b, _ := io.ReadAll(r)
	_ = r.Close()
	return string(b)
}

func captureStderr(t *testing.T, fn func()) string {
	t.Helper()

	orig := os.Stderr
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatalf("pipe: %v", err)
	}
	os.Stderr = w

	fn()

	_ = w.Close()
	os.Stderr = orig
	b, _ := io.ReadAll(r)
	_ = r.Close()
	return string(b)
}

func withStdin(t *testing.T, input string, fn func()) {
	t.Helper()

	orig := os.Stdin
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatalf("pipe: %v", err)
	}
	os.Stdin = r

	_, _ = io.WriteString(w, input)
	_ = w.Close()

	fn()

	_ = r.Close()
	os.Stdin = orig
}
