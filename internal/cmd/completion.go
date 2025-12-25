package cmd

import (
	"os"

	"github.com/spf13/cobra"
)

func newCompletionCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "completion [bash|zsh|fish|powershell]",
		Short: "Generate shell completion scripts",
		Long: `Generate shell completion scripts for gog.

To load completions:

Bash:
  $ source <(gog completion bash)
  # To load completions for each session, execute once:
  # Linux:
  $ gog completion bash > /etc/bash_completion.d/gog
  # macOS:
  $ gog completion bash > $(brew --prefix)/etc/bash_completion.d/gog

Zsh:
  # If shell completion is not already enabled in your environment,
  # you will need to enable it. You can execute the following once:
  $ echo "autoload -U compinit; compinit" >> ~/.zshrc
  # To load completions for each session, execute once:
  $ gog completion zsh > "${fpath[1]}/_gog"
  # You will need to start a new shell for this setup to take effect.

Fish:
  $ gog completion fish | source
  # To load completions for each session, execute once:
  $ gog completion fish > ~/.config/fish/completions/gog.fish

PowerShell:
  PS> gog completion powershell | Out-String | Invoke-Expression
  # To load completions for every new session, run:
  PS> gog completion powershell > gog.ps1
  # and source this file from your PowerShell profile.
`,
		DisableFlagsInUseLine: true,
		ValidArgs:             []string{"bash", "zsh", "fish", "powershell"},
		Args:                  cobra.MatchAll(cobra.ExactArgs(1), cobra.OnlyValidArgs),
		RunE: func(cmd *cobra.Command, args []string) error {
			switch args[0] {
			case "bash":
				return cmd.Root().GenBashCompletion(os.Stdout)
			case "zsh":
				return cmd.Root().GenZshCompletion(os.Stdout)
			case "fish":
				return cmd.Root().GenFishCompletion(os.Stdout, true)
			case "powershell":
				return cmd.Root().GenPowerShellCompletionWithDesc(os.Stdout)
			}
			return nil
		},
	}
	return cmd
}
