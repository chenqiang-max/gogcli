package googleapi

import (
	"context"
	"crypto/tls"
	"log/slog"
	"net/http"
	"time"

	"github.com/99designs/keyring"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"

	"github.com/steipete/gogcli/internal/config"
	"github.com/steipete/gogcli/internal/googleauth"
	"github.com/steipete/gogcli/internal/secrets"
)

const defaultHTTPTimeout = 30 * time.Second

var (
	readClientCredentials = config.ReadClientCredentials
	openSecretsStore      = secrets.OpenDefault
)

func tokenSourceForAccount(ctx context.Context, service googleauth.Service, email string) (oauth2.TokenSource, error) {
	creds, err := readClientCredentials()
	if err != nil {
		return nil, err
	}

	requiredScopes, err := googleauth.Scopes(service)
	if err != nil {
		return nil, err
	}

	return tokenSourceForAccountScopes(ctx, string(service), email, creds.ClientID, creds.ClientSecret, requiredScopes)
}

func tokenSourceForAccountScopes(ctx context.Context, serviceLabel string, email string, clientID string, clientSecret string, requiredScopes []string) (oauth2.TokenSource, error) {
	store, err := openSecretsStore()
	if err != nil {
		return nil, err
	}
	tok, err := store.GetToken(email)
	if err != nil {
		if err == keyring.ErrKeyNotFound {
			return nil, &AuthRequiredError{Service: serviceLabel, Email: email, Cause: err}
		}
		return nil, err
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
	slog.Debug("creating client options", "service", service, "email", email)

	ts, err := tokenSourceForAccount(ctx, service, email)
	if err != nil {
		return nil, err
	}
	baseTransport := &http.Transport{
		TLSClientConfig: &tls.Config{
			MinVersion: tls.VersionTLS12,
		},
	}
	// Wrap with retry logic for 429 and 5xx errors
	retryTransport := NewRetryTransport(&oauth2.Transport{
		Source: ts,
		Base:   baseTransport,
	})
	c := &http.Client{
		Transport: retryTransport,
		Timeout:   defaultHTTPTimeout,
	}

	slog.Debug("client options created successfully", "service", service, "email", email)
	return []option.ClientOption{option.WithHTTPClient(c)}, nil
}

func optionsForAccountScopes(ctx context.Context, serviceLabel string, email string, scopes []string) ([]option.ClientOption, error) {
	slog.Debug("creating client options with custom scopes", "serviceLabel", serviceLabel, "email", email)

	creds, err := readClientCredentials()
	if err != nil {
		return nil, err
	}
	ts, err := tokenSourceForAccountScopes(ctx, serviceLabel, email, creds.ClientID, creds.ClientSecret, scopes)
	if err != nil {
		return nil, err
	}
	baseTransport := &http.Transport{
		TLSClientConfig: &tls.Config{
			MinVersion: tls.VersionTLS12,
		},
	}
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
