package cmd

import (
	"testing"

	"github.com/spf13/cobra"
)

func TestMustMarkRequired(t *testing.T) {
	t.Run("valid flag", func(t *testing.T) {
		cmd := &cobra.Command{}
		cmd.Flags().String("test", "", "test flag")
		// Should not panic
		mustMarkRequired(cmd, "test")
	})

	t.Run("invalid flag panics", func(t *testing.T) {
		defer func() {
			if r := recover(); r == nil {
				t.Error("expected panic for non-existent flag")
			}
		}()
		cmd := &cobra.Command{}
		mustMarkRequired(cmd, "nonexistent")
	})
}

func TestValidateDate(t *testing.T) {
	tests := []struct {
		name    string
		input   string
		wantErr bool
	}{
		{"empty string", "", false},
		{"valid date", "2025-01-15", false},
		{"valid date leap year", "2024-02-29", false},
		{"invalid format", "01/15/2025", true},
		{"invalid format dashes", "2025-1-15", true},
		{"invalid date", "2025-13-01", true},
		{"invalid day", "2025-02-30", true},
		{"not a date", "not-a-date", true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := validateDate(tt.input)
			if (err != nil) != tt.wantErr {
				t.Errorf("validateDate(%q) error = %v, wantErr %v", tt.input, err, tt.wantErr)
			}
		})
	}
}

func TestValidateDateTime(t *testing.T) {
	tests := []struct {
		name    string
		input   string
		wantErr bool
	}{
		{"empty string", "", false},
		{"valid RFC3339", "2025-01-15T10:30:00Z", false},
		{"valid with timezone", "2025-01-15T10:30:00-05:00", false},
		{"valid with milliseconds", "2025-01-15T10:30:00.123Z", false},
		{"invalid format", "2025-01-15 10:30:00", true},
		{"date only", "2025-01-15", true},
		{"not a datetime", "not-a-datetime", true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := validateDateTime(tt.input)
			if (err != nil) != tt.wantErr {
				t.Errorf("validateDateTime(%q) error = %v, wantErr %v", tt.input, err, tt.wantErr)
			}
		})
	}
}

func TestValidateDateRange(t *testing.T) {
	tests := []struct {
		name    string
		from    string
		to      string
		wantErr bool
	}{
		{"both empty", "", "", false},
		{"from empty", "", "2025-01-15", false},
		{"to empty", "2025-01-15", "", false},
		{"valid range", "2025-01-01", "2025-01-31", false},
		{"same date", "2025-01-15", "2025-01-15", false},
		{"from after to", "2025-01-31", "2025-01-01", true},
		{"invalid from date", "2025-13-01", "2025-01-31", true},
		{"invalid to date", "2025-01-01", "2025-13-31", true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := validateDateRange(tt.from, tt.to)
			if (err != nil) != tt.wantErr {
				t.Errorf("validateDateRange(%q, %q) error = %v, wantErr %v", tt.from, tt.to, err, tt.wantErr)
			}
		})
	}
}

func TestValidateEmail(t *testing.T) {
	tests := []struct {
		name    string
		input   string
		wantErr bool
	}{
		{"empty string", "", false},
		{"valid simple email", "user@example.com", false},
		{"valid with subdomain", "user@mail.example.com", false},
		{"valid with plus", "user+tag@example.com", false},
		{"valid with display name", "User Name <user@example.com>", false},
		{"invalid no @", "userexample.com", true},
		{"invalid no domain", "user@", true},
		{"invalid no local", "@example.com", true},
		{"invalid spaces", "user @example.com", true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := validateEmail(tt.input)
			if (err != nil) != tt.wantErr {
				t.Errorf("validateEmail(%q) error = %v, wantErr %v", tt.input, err, tt.wantErr)
			}
		})
	}
}

func TestValidatePositiveInt(t *testing.T) {
	tests := []struct {
		name    string
		value   int64
		wantErr bool
	}{
		{"positive", 1, false},
		{"large positive", 1000000, false},
		{"zero", 0, true},
		{"negative", -1, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := validatePositiveInt(tt.value, "test")
			if (err != nil) != tt.wantErr {
				t.Errorf("validatePositiveInt(%d) error = %v, wantErr %v", tt.value, err, tt.wantErr)
			}
		})
	}
}

func TestConvertDateToRFC3339(t *testing.T) {
	tests := []struct {
		name    string
		input   string
		want    string
		wantErr bool
	}{
		{"valid date", "2025-01-15", "2025-01-15T00:00:00Z", false},
		{"leap year", "2024-02-29", "2024-02-29T00:00:00Z", false},
		{"invalid format", "01/15/2025", "", true},
		{"invalid date", "2025-13-01", "", true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := convertDateToRFC3339(tt.input)
			if (err != nil) != tt.wantErr {
				t.Errorf("convertDateToRFC3339(%q) error = %v, wantErr %v", tt.input, err, tt.wantErr)
				return
			}
			if !tt.wantErr && got != tt.want {
				t.Errorf("convertDateToRFC3339(%q) = %q, want %q", tt.input, got, tt.want)
			}
		})
	}
}
