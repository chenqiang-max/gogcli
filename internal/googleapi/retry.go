package googleapi

import (
	"context"
	"log/slog"
	"math/rand"
	"net/http"
	"time"

	"google.golang.org/api/googleapi"
)

const (
	// MaxRateLimitRetries is the maximum number of retries on 429 responses
	MaxRateLimitRetries = 3
	// RateLimitBaseDelay is the initial delay for rate limit exponential backoff
	RateLimitBaseDelay = 1 * time.Second
	// Max5xxRetries is the maximum retries for server errors
	Max5xxRetries = 1
	// ServerErrorRetryDelay is the delay before retrying on 5xx errors
	ServerErrorRetryDelay = 1 * time.Second
)

// RetryConfig configures retry behavior
type RetryConfig struct {
	MaxRateLimitRetries int
	Max5xxRetries       int
	BaseDelay           time.Duration
	CircuitBreaker      *CircuitBreaker
}

// DefaultRetryConfig returns the default retry configuration
func DefaultRetryConfig() *RetryConfig {
	return &RetryConfig{
		MaxRateLimitRetries: MaxRateLimitRetries,
		Max5xxRetries:       Max5xxRetries,
		BaseDelay:           RateLimitBaseDelay,
		CircuitBreaker:      NewCircuitBreaker(),
	}
}

// WithRetry wraps a Google API call with retry logic
func WithRetry[T any](ctx context.Context, cfg *RetryConfig, fn func() (T, error)) (T, error) {
	var zero T

	if cfg == nil {
		cfg = DefaultRetryConfig()
	}

	if cfg.CircuitBreaker != nil && cfg.CircuitBreaker.IsOpen() {
		return zero, &CircuitBreakerError{}
	}

	retries429 := 0
	retries5xx := 0

	for {
		result, err := fn()
		if err == nil {
			if cfg.CircuitBreaker != nil {
				cfg.CircuitBreaker.RecordSuccess()
			}
			return result, nil
		}

		// Check if it's a Google API error
		gerr, ok := err.(*googleapi.Error)
		if !ok {
			return zero, err
		}

		// 429 rate limit: exponential backoff with jitter
		if gerr.Code == http.StatusTooManyRequests {
			if retries429 >= cfg.MaxRateLimitRetries {
				return zero, &RateLimitError{Retries: retries429}
			}

			// Calculate backoff: 1s, 2s, 4s with jitter
			baseDelay := cfg.BaseDelay * time.Duration(1<<retries429)
			jitter := time.Duration(rand.Int63n(int64(baseDelay / 2)))
			delay := baseDelay + jitter

			// Check for Retry-After header hint in error message
			if retryAfter := parseRetryAfter(gerr); retryAfter > 0 {
				delay = retryAfter
			}

			slog.Info("rate limited, retrying", "delay", delay, "attempt", retries429+1, "max_retries", cfg.MaxRateLimitRetries)

			select {
			case <-time.After(delay):
			case <-ctx.Done():
				return zero, ctx.Err()
			}

			retries429++
			continue
		}

		// 5xx errors: retry once after delay
		if gerr.Code >= 500 {
			if cfg.CircuitBreaker != nil {
				cfg.CircuitBreaker.RecordFailure()
			}

			if retries5xx >= cfg.Max5xxRetries {
				return zero, err
			}

			slog.Info("server error, retrying", "status", gerr.Code, "attempt", retries5xx+1)

			select {
			case <-time.After(ServerErrorRetryDelay):
			case <-ctx.Done():
				return zero, ctx.Err()
			}

			retries5xx++
			continue
		}

		// Other errors: don't retry
		return zero, err
	}
}

// parseRetryAfter attempts to extract retry delay from a Google API error.
// Returns 0 as googleapi.Error doesn't expose HTTP headers.
// The transport layer (RetryTransport) handles Retry-After headers directly.
func parseRetryAfter(_ *googleapi.Error) time.Duration {
	return 0
}
