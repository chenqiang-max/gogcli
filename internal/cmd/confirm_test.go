package cmd

import (
	"context"
	"testing"

	"github.com/spf13/cobra"
	"github.com/steipete/gogcli/internal/ui"
)

func TestConfirmDestructive_NonInteractiveRequiresForce(t *testing.T) {
	cmd := &cobra.Command{}
	u, err := ui.New(ui.Options{Color: "never"})
	if err != nil {
		t.Fatalf("ui.New: %v", err)
	}
	cmd.SetContext(ui.WithUI(context.Background(), u))

	withStdin(t, "y\n", func() {
		flags := &rootFlags{Force: false, NoInput: false}
		err := confirmDestructive(cmd, flags, "delete something")
		if err == nil {
			t.Fatalf("expected error")
		}
		if ExitCode(err) != 2 {
			t.Fatalf("expected exit code 2, got %d (err=%v)", ExitCode(err), err)
		}
	})

	withStdin(t, "", func() {
		flags := &rootFlags{Force: true, NoInput: false}
		if err := confirmDestructive(cmd, flags, "delete something"); err != nil {
			t.Fatalf("expected no error, got %v", err)
		}
	})
}
