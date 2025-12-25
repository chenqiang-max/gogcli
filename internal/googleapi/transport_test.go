package googleapi

import (
	"context"
	"io"
	"net/http"
	"strings"
	"testing"
	"time"
)

// mockTransport implements http.RoundTripper for testing
type mockTransport struct {
	responses []*http.Response
	errors    []error
	calls     int
}

func (m *mockTransport) RoundTrip(req *http.Request) (*http.Response, error) {
	idx := m.calls
	m.calls++

	if idx < len(m.errors) && m.errors[idx] != nil {
		return nil, m.errors[idx]
	}
	if idx < len(m.responses) {
		return m.responses[idx], nil
	}
	return &http.Response{
		StatusCode: 200,
		Body:       io.NopCloser(strings.NewReader("")),
	}, nil
}

func TestRetryTransport_Success(t *testing.T) {
	mock := &mockTransport{
		responses: []*http.Response{
			{StatusCode: 200, Body: io.NopCloser(strings.NewReader("ok"))},
		},
	}

	rt := NewRetryTransport(mock)
	req, _ := http.NewRequest("GET", "https://example.com", nil)
	resp, err := rt.RoundTrip(req)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if resp.StatusCode != 200 {
		t.Errorf("expected 200, got %d", resp.StatusCode)
	}
	if mock.calls != 1 {
		t.Errorf("expected 1 call, got %d", mock.calls)
	}
}

func TestRetryTransport_RateLimit_Retry(t *testing.T) {
	mock := &mockTransport{
		responses: []*http.Response{
			{StatusCode: 429, Body: io.NopCloser(strings.NewReader("rate limited"))},
			{StatusCode: 200, Body: io.NopCloser(strings.NewReader("ok"))},
		},
	}

	rt := NewRetryTransport(mock)
	rt.BaseDelay = 10 * time.Millisecond // Speed up test

	req, _ := http.NewRequest("GET", "https://example.com", nil)
	resp, err := rt.RoundTrip(req)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if resp.StatusCode != 200 {
		t.Errorf("expected 200 after retry, got %d", resp.StatusCode)
	}
	if mock.calls != 2 {
		t.Errorf("expected 2 calls (1 retry), got %d", mock.calls)
	}
}

func TestRetryTransport_RateLimit_MaxRetries(t *testing.T) {
	// All responses are 429
	mock := &mockTransport{
		responses: []*http.Response{
			{StatusCode: 429, Body: io.NopCloser(strings.NewReader("rate limited"))},
			{StatusCode: 429, Body: io.NopCloser(strings.NewReader("rate limited"))},
			{StatusCode: 429, Body: io.NopCloser(strings.NewReader("rate limited"))},
			{StatusCode: 429, Body: io.NopCloser(strings.NewReader("rate limited"))},
		},
	}

	rt := NewRetryTransport(mock)
	rt.BaseDelay = 1 * time.Millisecond
	rt.MaxRetries429 = 2

	req, _ := http.NewRequest("GET", "https://example.com", nil)
	resp, err := rt.RoundTrip(req)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if resp.StatusCode != 429 {
		t.Errorf("expected 429 after max retries, got %d", resp.StatusCode)
	}
	// 1 initial + 2 retries = 3 total
	if mock.calls != 3 {
		t.Errorf("expected 3 calls, got %d", mock.calls)
	}
}

func TestRetryTransport_ServerError_Retry(t *testing.T) {
	mock := &mockTransport{
		responses: []*http.Response{
			{StatusCode: 503, Body: io.NopCloser(strings.NewReader("service unavailable"))},
			{StatusCode: 200, Body: io.NopCloser(strings.NewReader("ok"))},
		},
	}

	rt := NewRetryTransport(mock)

	req, _ := http.NewRequest("GET", "https://example.com", nil)
	resp, err := rt.RoundTrip(req)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if resp.StatusCode != 200 {
		t.Errorf("expected 200 after retry, got %d", resp.StatusCode)
	}
	if mock.calls != 2 {
		t.Errorf("expected 2 calls, got %d", mock.calls)
	}
}

func TestRetryTransport_ClientError_NoRetry(t *testing.T) {
	mock := &mockTransport{
		responses: []*http.Response{
			{StatusCode: 404, Body: io.NopCloser(strings.NewReader("not found"))},
		},
	}

	rt := NewRetryTransport(mock)

	req, _ := http.NewRequest("GET", "https://example.com", nil)
	resp, err := rt.RoundTrip(req)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if resp.StatusCode != 404 {
		t.Errorf("expected 404, got %d", resp.StatusCode)
	}
	if mock.calls != 1 {
		t.Errorf("expected 1 call (no retry for 4xx), got %d", mock.calls)
	}
}

func TestRetryTransport_ContextCanceled(t *testing.T) {
	mock := &mockTransport{
		responses: []*http.Response{
			{StatusCode: 429, Body: io.NopCloser(strings.NewReader("rate limited"))},
		},
	}

	rt := NewRetryTransport(mock)
	rt.BaseDelay = 1 * time.Second // Long delay

	ctx, cancel := context.WithCancel(context.Background())
	req, _ := http.NewRequestWithContext(ctx, "GET", "https://example.com", nil)

	// Cancel context during the backoff
	go func() {
		time.Sleep(10 * time.Millisecond)
		cancel()
	}()

	_, err := rt.RoundTrip(req)

	if err != context.Canceled {
		t.Errorf("expected context.Canceled, got %v", err)
	}
}

func TestRetryTransport_CircuitBreakerOpen(t *testing.T) {
	mock := &mockTransport{}

	rt := NewRetryTransport(mock)
	// Force circuit breaker open
	for i := 0; i < CircuitBreakerThreshold; i++ {
		rt.CircuitBreaker.RecordFailure()
	}

	req, _ := http.NewRequest("GET", "https://example.com", nil)
	_, err := rt.RoundTrip(req)

	if err == nil {
		t.Fatal("expected error when circuit breaker is open")
	}
	if !IsCircuitBreakerError(err) {
		t.Errorf("expected CircuitBreakerError, got %T", err)
	}
	if mock.calls != 0 {
		t.Errorf("expected 0 calls when circuit open, got %d", mock.calls)
	}
}

func TestRetryTransport_CircuitBreakerReset(t *testing.T) {
	mock := &mockTransport{
		responses: []*http.Response{
			{StatusCode: 200, Body: io.NopCloser(strings.NewReader("ok"))},
		},
	}

	rt := NewRetryTransport(mock)
	// Record failures but not enough to open
	for i := 0; i < CircuitBreakerThreshold-1; i++ {
		rt.CircuitBreaker.RecordFailure()
	}

	req, _ := http.NewRequest("GET", "https://example.com", nil)
	resp, err := rt.RoundTrip(req)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if resp.StatusCode != 200 {
		t.Errorf("expected 200, got %d", resp.StatusCode)
	}

	// After success, failures should be reset
	if rt.CircuitBreaker.failures != 0 {
		t.Errorf("expected failures reset to 0, got %d", rt.CircuitBreaker.failures)
	}
}

func TestRetryTransport_RetryAfterHeader(t *testing.T) {
	mock := &mockTransport{
		responses: []*http.Response{
			{
				StatusCode: 429,
				Header:     http.Header{"Retry-After": []string{"1"}},
				Body:       io.NopCloser(strings.NewReader("rate limited")),
			},
			{StatusCode: 200, Body: io.NopCloser(strings.NewReader("ok"))},
		},
	}

	rt := NewRetryTransport(mock)
	rt.BaseDelay = 1 * time.Hour // Would be very long without Retry-After

	start := time.Now()
	req, _ := http.NewRequest("GET", "https://example.com", nil)
	resp, err := rt.RoundTrip(req)
	elapsed := time.Since(start)

	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if resp.StatusCode != 200 {
		t.Errorf("expected 200, got %d", resp.StatusCode)
	}
	// Should have waited ~1 second based on Retry-After header
	if elapsed < 900*time.Millisecond || elapsed > 2*time.Second {
		t.Errorf("expected ~1s delay from Retry-After, got %v", elapsed)
	}
}

func TestRetryTransport_WithRequestBody(t *testing.T) {
	mock := &mockTransport{
		responses: []*http.Response{
			{StatusCode: 429, Body: io.NopCloser(strings.NewReader("rate limited"))},
			{StatusCode: 200, Body: io.NopCloser(strings.NewReader("ok"))},
		},
	}

	rt := NewRetryTransport(mock)
	rt.BaseDelay = 1 * time.Millisecond

	body := strings.NewReader("request body")
	req, _ := http.NewRequest("POST", "https://example.com", body)
	resp, err := rt.RoundTrip(req)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if resp.StatusCode != 200 {
		t.Errorf("expected 200 after retry, got %d", resp.StatusCode)
	}
	if mock.calls != 2 {
		t.Errorf("expected 2 calls, got %d", mock.calls)
	}
}
