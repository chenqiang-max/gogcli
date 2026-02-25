package googleapi

import (
	"context"
	"crypto/tls"
	"errors"
	"fmt"
	"log/slog"
	"net/http"
	"time"

	"github.com/99designs/keyring"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"

	"github.com/steipete/gogcli/internal/authclient"
	"github.com/steipete/gogcli/internal/config"
	"github.com/steipete/gogcli/internal/googleauth"
	"github.com/steipete/gogcli/internal/secrets"
)

const defaultHTTPTimeout = 30 * time.Second

var (
	readClientCredentials   = config.ReadClientCredentialsFor
	openSecretsStore        = secrets.OpenDefault
	readExternalTokenConfig = config.ReadExternalTokenConfig
)

// getExternalToken returns token from external config file for the given service
func getExternalToken(service googleauth.Service) string {
	extConfig, err := readExternalTokenConfig()
	if err != nil {
		slog.Debug("failed to read external token config", "error", err)
		return ""
	}

	// Map service to external token config field
	switch service {
	case googleauth.ServiceDrive:
		return extConfig.GetGDriveToken()
	case googleauth.ServiceGmail:
		return extConfig.GetGmailToken()
	case googleauth.ServiceCalendar:
		return extConfig.GetCalendarToken()
	default:
		return ""
	}
}

// getExternalTokenByLabel returns token from external config by service label string
func getExternalTokenByLabel(extConfig *config.ExternalTokenConfig, serviceLabel string) string {
	if extConfig == nil {
		return ""
	}

	// Map service label to external token config field
	switch serviceLabel {
	case "drive":
		return extConfig.GetGDriveToken()
	case "gmail":
		return extConfig.GetGmailToken()
	case "calendar":
		return extConfig.GetCalendarToken()
	default:
		return ""
	}
}

func tokenSourceForAccount(ctx context.Context, service googleauth.Service, email string) (oauth2.TokenSource, error) {
	// Check for external token first (from .manus-gogcli.conf)
	if extToken := getExternalToken(service); extToken != "" {
		slog.Debug("using external token", "service", service)
		return oauth2.StaticTokenSource(&oauth2.Token{
			AccessToken: extToken,
			TokenType:   "Bearer",
		}), nil
	}

	client, err := authclient.ResolveClient(ctx, email)
	if err != nil {
		return nil, fmt.Errorf("resolve client: %w", err)
	}

	creds, err := readClientCredentials(client)
	if err != nil {
		return nil, fmt.Errorf("read credentials: %w", err)
	}

	var requiredScopes []string

	if scopes, err := googleauth.Scopes(service); err != nil {
		return nil, fmt.Errorf("resolve scopes: %w", err)
	} else {
		requiredScopes = scopes
	}

	return tokenSourceForAccountScopes(ctx, string(service), email, client, creds.ClientID, creds.ClientSecret, requiredScopes)
}

func tokenSourceForAccountScopes(ctx context.Context, serviceLabel string, email string, client string, clientID string, clientSecret string, requiredScopes []string) (oauth2.TokenSource, error) {
	var store secrets.Store

	if s, err := openSecretsStore(); err != nil {
		return nil, fmt.Errorf("open secrets store: %w", err)
	} else {
		store = s
	}

	var tok secrets.Token

	if t, err := store.GetToken(client, email); err != nil {
		if errors.Is(err, keyring.ErrKeyNotFound) {
			return nil, &AuthRequiredError{Service: serviceLabel, Email: email, Client: client, Cause: err}
		}

		return nil, fmt.Errorf("get token for %s: %w", email, err)
	} else {
		tok = t
	}

	// If access_token is available, use it directly (skip OAuth refresh flow)
	if tok.AccessToken != "" {
		return oauth2.StaticTokenSource(&oauth2.Token{
			AccessToken: tok.AccessToken,
			TokenType:   "Bearer",
		}), nil
	}

	cfg := oauth2.Config{
		ClientID:     clientID,
		ClientSecret: clientSecret,
		Endpoint:     google.Endpoint,
		Scopes:       requiredScopes,
	}

	// Ensure refresh-token exchanges don't hang forever.
	ctx = context.WithValue(ctx, oauth2.HTTPClient, &http.Client{Timeout: defaultHTTPTimeout})

	return cfg.TokenSource(ctx, &oauth2.Token{RefreshToken: tok.RefreshToken}), nil
}

func optionsForAccount(ctx context.Context, service googleauth.Service, email string) ([]option.ClientOption, error) {
	scopes, err := googleauth.Scopes(service)
	if err != nil {
		return nil, fmt.Errorf("resolve scopes: %w", err)
	}

	return optionsForAccountScopes(ctx, string(service), email, scopes)
}

func optionsForAccountScopes(ctx context.Context, serviceLabel string, email string, scopes []string) ([]option.ClientOption, error) {
	slog.Debug("creating client options with custom scopes", "serviceLabel", serviceLabel, "email", email)

	var creds config.ClientCredentials

	var ts oauth2.TokenSource

	if serviceAccountTS, saPath, ok, err := tokenSourceForServiceAccountScopes(ctx, email, scopes); err != nil {
		return nil, fmt.Errorf("service account token source: %w", err)
	} else if ok {
		slog.Debug("using service account credentials", "email", email, "path", saPath)
		ts = serviceAccountTS
	} else {
		client, err := authclient.ResolveClient(ctx, email)
		if err != nil {
			return nil, fmt.Errorf("resolve client: %w", err)
		}

		// Check for external token first (from .manus-gogcli.conf)
		extConfig, _ := readExternalTokenConfig()
		extToken := getExternalTokenByLabel(extConfig, serviceLabel)
		if extToken != "" {
			slog.Debug("using external token for service", "service", serviceLabel)
			ts = oauth2.StaticTokenSource(&oauth2.Token{
				AccessToken: extToken,
				TokenType:   "Bearer",
			})
		} else {
			// Try to use tokenSourceForAccountScopes, which handles both access_token and refresh_token
			// It will try access_token first (if available), and fall back to refresh_token flow
			if tokenSource, err := tokenSourceForAccountScopes(ctx, serviceLabel, email, client, "", "", scopes); err == nil {
				ts = tokenSource
			} else {
				// If access_token or refresh_token are not available, try loading credentials
				if c, err := readClientCredentials(client); err != nil {
					return nil, fmt.Errorf("read credentials: %w", err)
				} else {
					creds = c
				}

				if tokenSource, err := tokenSourceForAccountScopes(ctx, serviceLabel, email, client, creds.ClientID, creds.ClientSecret, scopes); err != nil {
					return nil, fmt.Errorf("token source: %w", err)
				} else {
					ts = tokenSource
				}
			}
		}
	}
	baseTransport := newBaseTransport()
	// Wrap with retry logic for 429 and 5xx errors
	retryTransport := NewRetryTransport(&oauth2.Transport{
		Source: ts,
		Base:   baseTransport,
	})
	c := &http.Client{
		Transport: retryTransport,
		Timeout:   defaultHTTPTimeout,
	}

	slog.Debug("client options with custom scopes created successfully", "serviceLabel", serviceLabel, "email", email)

	return []option.ClientOption{option.WithHTTPClient(c)}, nil
}

func newBaseTransport() *http.Transport {
	defaultTransport, ok := http.DefaultTransport.(*http.Transport)
	if !ok || defaultTransport == nil {
		return &http.Transport{
			Proxy: http.ProxyFromEnvironment,
			TLSClientConfig: &tls.Config{
				MinVersion: tls.VersionTLS12,
			},
		}
	}

	// Clone() deep-copies TLSClientConfig, so no additional clone needed.
	transport := defaultTransport.Clone()
	if transport.TLSClientConfig == nil {
		transport.TLSClientConfig = &tls.Config{MinVersion: tls.VersionTLS12}
		return transport
	}

	if transport.TLSClientConfig.MinVersion < tls.VersionTLS12 {
		transport.TLSClientConfig.MinVersion = tls.VersionTLS12
	}

	return transport
}
