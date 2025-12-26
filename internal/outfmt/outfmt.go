package outfmt

import (
	"context"
	"encoding/json"
	"io"
	"os"
	"strings"
)

type Mode struct {
	JSON  bool
	Plain bool
}

type ParseError struct{ msg string }

func (e *ParseError) Error() string { return e.msg }

func FromFlags(jsonOut bool, plainOut bool) (Mode, error) {
	if jsonOut && plainOut {
		return Mode{}, &ParseError{msg: "invalid output mode (cannot combine --json and --plain)"}
	}
	return Mode{JSON: jsonOut, Plain: plainOut}, nil
}

func FromEnv() Mode {
	return Mode{
		JSON:  envBool("GOG_JSON"),
		Plain: envBool("GOG_PLAIN"),
	}
}

type ctxKey struct{}

func WithMode(ctx context.Context, mode Mode) context.Context {
	return context.WithValue(ctx, ctxKey{}, mode)
}

func FromContext(ctx context.Context) Mode {
	if v := ctx.Value(ctxKey{}); v != nil {
		if m, ok := v.(Mode); ok {
			return m
		}
	}
	return Mode{}
}

func IsJSON(ctx context.Context) bool  { return FromContext(ctx).JSON }
func IsPlain(ctx context.Context) bool { return FromContext(ctx).Plain }

func WriteJSON(w io.Writer, v any) error {
	enc := json.NewEncoder(w)
	enc.SetEscapeHTML(false)
	enc.SetIndent("", "  ")
	return enc.Encode(v)
}

func envBool(key string) bool {
	v := strings.TrimSpace(strings.ToLower(os.Getenv(key)))
	switch v {
	case "1", "true", "yes", "y", "on":
		return true
	default:
		return false
	}
}
