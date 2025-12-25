package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
	"github.com/steipete/gogcli/internal/outfmt"
)

var (
	// Version is the semantic version (set via ldflags)
	Version = "dev"
	// Commit is the git commit hash (set via ldflags)
	Commit = "unknown"
	// BuildDate is the build timestamp (set via ldflags)
	BuildDate = "unknown"
)

type versionInfo struct {
	Version   string `json:"version"`
	Commit    string `json:"commit"`
	BuildDate string `json:"build_date"`
}

func newVersionCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "version",
		Short: "Print version information",
		Long:  "Display version, commit hash, and build date for this binary.",
		RunE: func(cmd *cobra.Command, args []string) error {
			if outfmt.IsJSON(cmd.Context()) {
				info := versionInfo{
					Version:   Version,
					Commit:    Commit,
					BuildDate: BuildDate,
				}
				return outfmt.WriteJSON(cmd.OutOrStdout(), info)
			}

			fmt.Fprintf(cmd.OutOrStdout(), "gog %s\n", Version)
			fmt.Fprintf(cmd.OutOrStdout(), "  commit:     %s\n", Commit)
			fmt.Fprintf(cmd.OutOrStdout(), "  build date: %s\n", BuildDate)
			return nil
		},
	}
}
