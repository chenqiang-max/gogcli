package cmd

import (
	"context"
	"errors"
	"net"
	"net/http"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/steipete/gogcli/internal/outfmt"
	"github.com/steipete/gogcli/internal/ui"
	"google.golang.org/api/gmail/v1"
	"google.golang.org/api/idtoken"
)

var newOIDCValidator = idtoken.NewValidator

func newGmailWatchCmd(flags *rootFlags) *cobra.Command {
	cmd := &cobra.Command{
		Use:   "watch",
		Short: "Watch Gmail via Pub/Sub push",
	}

	cmd.AddCommand(newGmailWatchStartCmd(flags))
	cmd.AddCommand(newGmailWatchStatusCmd(flags))
	cmd.AddCommand(newGmailWatchRenewCmd(flags))
	cmd.AddCommand(newGmailWatchStopCmd(flags))
	cmd.AddCommand(newGmailWatchServeCmd(flags))
	return cmd
}

func newGmailWatchStartCmd(flags *rootFlags) *cobra.Command {
	var topic string
	var labels []string
	var ttlRaw string
	var hookURL string
	var hookToken string
	var includeBody bool
	var maxBytes int

	cmd := &cobra.Command{
		Use:   "start",
		Short: "Start Gmail watch for Pub/Sub",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			account, err := requireAccount(flags)
			if err != nil {
				return err
			}
			if strings.TrimSpace(topic) == "" {
				return usage("--topic is required")
			}
			ttl, err := parseDurationSeconds(ttlRaw)
			if err != nil {
				return err
			}
			maxChanged := cmd.Flags().Changed("max-bytes")
			hook, err := hookFromFlags(hookURL, hookToken, includeBody, maxBytes, maxChanged, false)
			if err != nil {
				return err
			}

			svc, err := newGmailService(cmd.Context(), account)
			if err != nil {
				return err
			}
			labelIDs, err := resolveLabelIDsWithService(svc, labels)
			if err != nil {
				return err
			}

			resp, err := requestGmailWatch(cmd.Context(), svc, topic, labelIDs)
			if err != nil {
				return err
			}
			state, err := buildWatchState(account, topic, labelIDs, resp, ttl, hook)
			if err != nil {
				return err
			}

			store, err := newGmailWatchStore(account)
			if err != nil {
				return err
			}
			if err := store.Update(func(s *gmailWatchState) error {
				*s = state
				return nil
			}); err != nil {
				return err
			}

			return writeWatchState(cmd.Context(), state)
		},
	}

	cmd.Flags().StringVar(&topic, "topic", "", "Pub/Sub topic (projects/.../topics/...)")
	cmd.Flags().StringSliceVar(&labels, "label", nil, "Label IDs or names (repeatable, comma-separated)")
	cmd.Flags().StringVar(&ttlRaw, "ttl", "", "Renew after duration (seconds or Go duration)")
	cmd.Flags().StringVar(&hookURL, "hook-url", "", "Webhook URL to forward messages")
	cmd.Flags().StringVar(&hookToken, "hook-token", "", "Webhook bearer token")
	cmd.Flags().BoolVar(&includeBody, "include-body", false, "Include text/plain body in hook payload")
	cmd.Flags().IntVar(&maxBytes, "max-bytes", defaultHookMaxBytes, "Max bytes of body to include")
	return cmd
}

func newGmailWatchStatusCmd(flags *rootFlags) *cobra.Command {
	return &cobra.Command{
		Use:   "status",
		Short: "Show stored watch state",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			account, err := requireAccount(flags)
			if err != nil {
				return err
			}
			store, err := loadGmailWatchStore(account)
			if err != nil {
				return err
			}
			return writeWatchState(cmd.Context(), store.Get())
		},
	}
}

func newGmailWatchRenewCmd(flags *rootFlags) *cobra.Command {
	var ttlRaw string

	cmd := &cobra.Command{
		Use:   "renew",
		Short: "Renew Gmail watch using stored config",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			account, err := requireAccount(flags)
			if err != nil {
				return err
			}
			store, err := loadGmailWatchStore(account)
			if err != nil {
				return err
			}
			state := store.Get()
			if strings.TrimSpace(state.Topic) == "" {
				return errors.New("stored watch state missing topic")
			}

			ttl, err := parseDurationSeconds(ttlRaw)
			if err != nil {
				return err
			}

			svc, err := newGmailService(cmd.Context(), account)
			if err != nil {
				return err
			}
			resp, err := requestGmailWatch(cmd.Context(), svc, state.Topic, state.Labels)
			if err != nil {
				return err
			}
			updated, err := buildWatchState(account, state.Topic, state.Labels, resp, ttl, state.Hook)
			if err != nil {
				return err
			}
			if ttl == 0 {
				updated.RenewAfterMs = state.RenewAfterMs
			}

			if err := store.Update(func(s *gmailWatchState) error {
				*s = updated
				return nil
			}); err != nil {
				return err
			}

			return writeWatchState(cmd.Context(), updated)
		},
	}

	cmd.Flags().StringVar(&ttlRaw, "ttl", "", "Renew after duration (seconds or Go duration)")
	return cmd
}

func newGmailWatchStopCmd(flags *rootFlags) *cobra.Command {
	return &cobra.Command{
		Use:   "stop",
		Short: "Stop Gmail watch and clear stored state",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			u := ui.FromContext(cmd.Context())
			account, err := requireAccount(flags)
			if err != nil {
				return err
			}

			if err := confirmDestructive(cmd, flags, "stop gmail watch and clear stored state"); err != nil {
				return err
			}

			svc, err := newGmailService(cmd.Context(), account)
			if err != nil {
				return err
			}
			if stopErr := svc.Users.Stop("me").Do(); stopErr != nil {
				return stopErr
			}
			store, err := newGmailWatchStore(account)
			if err == nil && store.path != "" {
				_ = os.Remove(store.path)
			}
			if outfmt.IsJSON(cmd.Context()) {
				return outfmt.WriteJSON(os.Stdout, map[string]any{"stopped": true})
			}
			u.Out().Printf("stopped\ttrue")
			return nil
		},
	}
}

func newGmailWatchServeCmd(flags *rootFlags) *cobra.Command {
	var bind string
	var port int
	var path string
	var verifyOIDC bool
	var oidcEmail string
	var oidcAudience string
	var sharedToken string
	var hookURL string
	var hookToken string
	var includeBody bool
	var maxBytes int
	var saveHook bool

	cmd := &cobra.Command{
		Use:   "serve",
		Short: "Run Pub/Sub push handler",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			u := ui.FromContext(cmd.Context())
			account, err := requireAccount(flags)
			if err != nil {
				return err
			}
			if !strings.HasPrefix(path, "/") {
				return usage("--path must start with '/'")
			}
			if port <= 0 {
				return usage("--port must be > 0")
			}
			if !verifyOIDC && sharedToken == "" && !isLoopbackHost(bind) {
				return usage("--verify-oidc or --token required when binding non-loopback")
			}
			if oidcEmail != "" && !verifyOIDC {
				return usage("--oidc-email requires --verify-oidc")
			}
			if oidcAudience != "" && !verifyOIDC {
				return usage("--oidc-audience requires --verify-oidc")
			}

			store, err := loadGmailWatchStore(account)
			if err != nil {
				return err
			}
			state := store.Get()

			if hookURL == "" && state.Hook != nil {
				hookURL = state.Hook.URL
				if !cmd.Flags().Changed("hook-token") {
					hookToken = state.Hook.Token
				}
				if !cmd.Flags().Changed("include-body") {
					includeBody = state.Hook.IncludeBody
				}
				if !cmd.Flags().Changed("max-bytes") && state.Hook.MaxBytes > 0 {
					maxBytes = state.Hook.MaxBytes
				}
			}

			maxChanged := cmd.Flags().Changed("max-bytes")
			hook, err := hookFromFlags(hookURL, hookToken, includeBody, maxBytes, maxChanged, true)
			if err != nil {
				return err
			}
			if saveHook && hook != nil {
				if updateErr := store.Update(func(s *gmailWatchState) error {
					s.Hook = hook
					s.UpdatedAtMs = time.Now().UnixMilli()
					return nil
				}); updateErr != nil {
					return updateErr
				}
			}

			validator := (*idtoken.Validator)(nil)
			if verifyOIDC {
				validator, err = newOIDCValidator(cmd.Context())
				if err != nil {
					return err
				}
			}

			cfg := gmailWatchServeConfig{
				Account:      account,
				Bind:         bind,
				Port:         port,
				Path:         path,
				VerifyOIDC:   verifyOIDC,
				OIDCEmail:    oidcEmail,
				OIDCAudience: oidcAudience,
				SharedToken:  sharedToken,
				HookTimeout:  defaultHookRequestTimeoutSec * time.Second,
				HistoryMax:   defaultHistoryMaxResults,
				ResyncMax:    defaultHistoryResyncMax,
				AllowNoHook:  hook == nil,
				IncludeBody:  includeBody,
				MaxBodyBytes: maxBytes,
			}
			if hook != nil {
				cfg.HookURL = hook.URL
				cfg.HookToken = hook.Token
				cfg.IncludeBody = hook.IncludeBody
				cfg.MaxBodyBytes = hook.MaxBytes
			}

			if cfg.MaxBodyBytes <= 0 {
				cfg.MaxBodyBytes = defaultHookMaxBytes
			}

			hookClient := &http.Client{Timeout: cfg.HookTimeout}
			server := &gmailWatchServer{
				cfg:        cfg,
				store:      store,
				validator:  validator,
				newService: newGmailService,
				hookClient: hookClient,
				logf:       u.Err().Printf,
				warnf:      u.Err().Printf,
			}

			addr := net.JoinHostPort(bind, strconv.Itoa(port))
			u.Err().Printf("watch: listening on %s%s", addr, path)

			httpServer := &http.Server{
				Addr:              addr,
				Handler:           server,
				ReadHeaderTimeout: 5 * time.Second,
			}
			return httpServer.ListenAndServe()
		},
	}

	cmd.Flags().StringVar(&bind, "bind", "127.0.0.1", "Bind address")
	cmd.Flags().IntVar(&port, "port", defaultWatchPort, "Listen port")
	cmd.Flags().StringVar(&path, "path", defaultWatchPath, "Push handler path")
	cmd.Flags().BoolVar(&verifyOIDC, "verify-oidc", false, "Verify Pub/Sub OIDC tokens")
	cmd.Flags().StringVar(&oidcEmail, "oidc-email", "", "Expected service account email")
	cmd.Flags().StringVar(&oidcAudience, "oidc-audience", "", "Expected OIDC audience")
	cmd.Flags().StringVar(&sharedToken, "token", "", "Shared token for x-gog-token or ?token=")
	cmd.Flags().StringVar(&hookURL, "hook-url", "", "Webhook URL to forward messages")
	cmd.Flags().StringVar(&hookToken, "hook-token", "", "Webhook bearer token")
	cmd.Flags().BoolVar(&includeBody, "include-body", false, "Include text/plain body in hook payload")
	cmd.Flags().IntVar(&maxBytes, "max-bytes", defaultHookMaxBytes, "Max bytes of body to include")
	cmd.Flags().BoolVar(&saveHook, "save-hook", false, "Persist hook settings to watch state")
	return cmd
}

func writeWatchState(ctx context.Context, state gmailWatchState) error {
	if outfmt.IsJSON(ctx) {
		return outfmt.WriteJSON(os.Stdout, map[string]any{"watch": state})
	}
	u := ui.FromContext(ctx)
	u.Out().Printf("account\t%s", state.Account)
	u.Out().Printf("topic\t%s", state.Topic)
	if len(state.Labels) > 0 {
		u.Out().Printf("labels\t%s", strings.Join(state.Labels, ","))
	}
	u.Out().Printf("history_id\t%s", state.HistoryID)
	if state.ExpirationMs > 0 {
		u.Out().Printf("expiration\t%s", formatUnixMillis(state.ExpirationMs))
	}
	if state.ProviderExpirationMs > 0 {
		u.Out().Printf("provider_expiration\t%s", formatUnixMillis(state.ProviderExpirationMs))
	}
	if state.RenewAfterMs > 0 {
		u.Out().Printf("renew_after\t%s", formatUnixMillis(state.RenewAfterMs))
	}
	if state.UpdatedAtMs > 0 {
		u.Out().Printf("updated_at\t%s", formatUnixMillis(state.UpdatedAtMs))
	}
	if state.Hook != nil {
		u.Out().Printf("hook_url\t%s", state.Hook.URL)
		if state.Hook.IncludeBody {
			u.Out().Printf("hook_include_body\ttrue")
		}
		if state.Hook.MaxBytes > 0 {
			u.Out().Printf("hook_max_bytes\t%d", state.Hook.MaxBytes)
		}
		if state.Hook.Token != "" {
			u.Out().Printf("hook_token\t%s", state.Hook.Token)
		}
	}
	if state.LastDeliveryStatus != "" {
		u.Out().Printf("last_delivery_status\t%s", state.LastDeliveryStatus)
	}
	if state.LastDeliveryAtMs > 0 {
		u.Out().Printf("last_delivery_at\t%s", formatUnixMillis(state.LastDeliveryAtMs))
	}
	if state.LastDeliveryStatusNote != "" {
		u.Out().Printf("last_delivery_note\t%s", state.LastDeliveryStatusNote)
	}
	return nil
}

func buildWatchState(account, topic string, labels []string, resp *gmail.WatchResponse, ttl time.Duration, hook *gmailWatchHook) (gmailWatchState, error) {
	if resp == nil {
		return gmailWatchState{}, errors.New("watch response missing")
	}
	historyID := formatHistoryID(resp.HistoryId)
	if historyID == "" {
		return gmailWatchState{}, errors.New("watch response missing historyId")
	}
	now := time.Now()
	state := gmailWatchState{
		Account:              account,
		Topic:                topic,
		Labels:               labels,
		HistoryID:            historyID,
		ExpirationMs:         resp.Expiration,
		ProviderExpirationMs: resp.Expiration,
		UpdatedAtMs:          now.UnixMilli(),
		Hook:                 hook,
	}
	if ttl > 0 {
		state.RenewAfterMs = now.Add(ttl).UnixMilli()
	}
	return state, nil
}

func requestGmailWatch(ctx context.Context, svc *gmail.Service, topic string, labelIDs []string) (*gmail.WatchResponse, error) {
	req := &gmail.WatchRequest{TopicName: topic}
	if len(labelIDs) > 0 {
		req.LabelIds = labelIDs
	}
	return svc.Users.Watch("me", req).Context(ctx).Do()
}

func hookFromFlags(url, token string, includeBody bool, maxBytes int, maxBytesChanged bool, allowNoHook bool) (*gmailWatchHook, error) {
	if strings.TrimSpace(url) == "" {
		if token != "" {
			return nil, usage("--hook-url required when using --hook-token")
		}
		if !allowNoHook && (includeBody || maxBytesChanged) {
			return nil, usage("--hook-url required when setting hook options")
		}
		return nil, nil
	}
	if maxBytes <= 0 {
		if includeBody {
			maxBytes = defaultHookMaxBytes
		} else if maxBytesChanged {
			return nil, usage("--max-bytes must be > 0")
		}
	}
	return &gmailWatchHook{
		URL:         url,
		Token:       token,
		IncludeBody: includeBody,
		MaxBytes:    maxBytes,
	}, nil
}

func isLoopbackHost(host string) bool {
	trimmed := strings.TrimSpace(host)
	if trimmed == "" {
		return true
	}
	if strings.EqualFold(trimmed, "localhost") {
		return true
	}
	trimmed = strings.TrimPrefix(trimmed, "[")
	trimmed = strings.TrimSuffix(trimmed, "]")
	ip := net.ParseIP(trimmed)
	if ip == nil {
		return false
	}
	return ip.IsLoopback()
}
