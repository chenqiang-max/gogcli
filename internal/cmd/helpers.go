package cmd

import (
	"fmt"
	"net/mail"
	"time"

	"github.com/spf13/cobra"
)

// mustMarkRequired marks a flag as required, panicking on error.
// Use for flags that are definitely defined - panics indicate programmer error.
func mustMarkRequired(cmd *cobra.Command, name string) {
	if err := cmd.MarkFlagRequired(name); err != nil {
		panic(fmt.Sprintf("flag %q not defined: %v", name, err))
	}
}

// validateDate validates that a date string is in YYYY-MM-DD format
func validateDate(dateStr string) error {
	if dateStr == "" {
		return nil // empty is valid (optional parameter)
	}
	_, err := time.Parse("2006-01-02", dateStr)
	if err != nil {
		return fmt.Errorf("invalid date format: expected YYYY-MM-DD, got %q", dateStr)
	}
	return nil
}

// validateDateTime validates that a string is in RFC3339 format
func validateDateTime(dateTimeStr string) error {
	if dateTimeStr == "" {
		return nil
	}
	_, err := time.Parse(time.RFC3339, dateTimeStr)
	if err != nil {
		return fmt.Errorf("invalid datetime format: expected RFC3339 (e.g., 2006-01-02T15:04:05Z), got %q", dateTimeStr)
	}
	return nil
}

// validateDateRange validates that from date is before to date when both are provided
func validateDateRange(from, to string) error {
	if from == "" || to == "" {
		return nil // only validate if both are provided
	}

	fromTime, err := time.Parse("2006-01-02", from)
	if err != nil {
		return fmt.Errorf("invalid from date: expected YYYY-MM-DD, got %q", from)
	}

	toTime, err := time.Parse("2006-01-02", to)
	if err != nil {
		return fmt.Errorf("invalid to date: expected YYYY-MM-DD, got %q", to)
	}

	if fromTime.After(toTime) {
		return fmt.Errorf("from date (%s) must be before or equal to to date (%s)", from, to)
	}

	return nil
}

// validateEmail validates that a string is a valid email address
func validateEmail(email string) error {
	if email == "" {
		return nil
	}
	_, err := mail.ParseAddress(email)
	if err != nil {
		return fmt.Errorf("invalid email address: %q", email)
	}
	return nil
}

// validatePositiveInt validates that an integer is positive
func validatePositiveInt(value int64, name string) error {
	if value <= 0 {
		return fmt.Errorf("%s must be positive, got %d", name, value)
	}
	return nil
}

// convertDateToRFC3339 converts a date string in YYYY-MM-DD format to RFC3339 format
// with time set to 00:00:00 UTC
func convertDateToRFC3339(dateStr string) (string, error) {
	t, err := time.Parse("2006-01-02", dateStr)
	if err != nil {
		return "", fmt.Errorf("expected format YYYY-MM-DD, got %q", dateStr)
	}
	return t.UTC().Format(time.RFC3339), nil
}

// Note: splitCSV and orEmpty are already defined in calendar.go and used across commands.
// They should remain in their current location to avoid import cycles.
