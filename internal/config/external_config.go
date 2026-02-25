package config

import (
	"fmt"
	"os"
	"strings"
)

const (
	ExternalTokenConfigPath = "/home/ubuntu/.manus-gogcli.conf"
)

// ExternalTokenConfig holds tokens from external configuration file (.manus-gogcli.conf)
type ExternalTokenConfig struct {
	GDriveEmail   string // email for Google Drive (optional)
	GDriveToken   string // token for Google Drive
	GmailEmail    string // email for Gmail (optional)
	GmailToken    string // token for Gmail (future use)
	CalendarEmail string // email for Google Calendar (optional)
	CalendarToken string // token for Google Calendar (future use)
}

// ReadExternalTokenConfig reads tokens from ~/.manus-gogcli.conf
func ReadExternalTokenConfig() (*ExternalTokenConfig, error) {
	configPath := ExternalTokenConfigPath

	data, err := os.ReadFile(configPath)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, nil // Config file doesn't exist, return nil (not an error)
		}
		return nil, fmt.Errorf("read config file: %w", err)
	}

	return parseExternalConfig(string(data))
}

func parseExternalConfig(content string) (*ExternalTokenConfig, error) {
	config := &ExternalTokenConfig{}
	lines := strings.Split(content, "\n")

	var currentSection string
	for _, line := range lines {
		line = strings.TrimSpace(line)

		// Skip empty lines and comments
		if line == "" || strings.HasPrefix(line, "#") || strings.HasPrefix(line, ";") {
			continue
		}

		// Parse section headers like [gdrive], [gmail], [calendar]
		if strings.HasPrefix(line, "[") && strings.HasSuffix(line, "]") {
			currentSection = strings.ToLower(strings.TrimSuffix(strings.TrimPrefix(line, "["), "]"))
			continue
		}

		// Parse key=value pairs
		if !strings.Contains(line, "=") {
			continue
		}

		parts := strings.SplitN(line, "=", 2)
		if len(parts) != 2 {
			continue
		}

		key := strings.TrimSpace(parts[0])
		value := strings.TrimSpace(parts[1])

		// Remove quotes if present
		if (strings.HasPrefix(value, "\"") && strings.HasSuffix(value, "\"")) ||
			(strings.HasPrefix(value, "'") && strings.HasSuffix(value, "'")) {
			value = value[1 : len(value)-1]
		}

		// Store tokens and emails by section
		switch key {
		case "token":
			switch currentSection {
			case "gdrive":
				config.GDriveToken = value
			case "gmail":
				config.GmailToken = value
			case "calendar":
				config.CalendarToken = value
			}
		case "email":
			switch currentSection {
			case "gdrive":
				config.GDriveEmail = value
			case "gmail":
				config.GmailEmail = value
			case "calendar":
				config.CalendarEmail = value
			}
		}
	}

	return config, nil
}

// GetGDriveEmail returns the Google Drive email from external config
func (c *ExternalTokenConfig) GetGDriveEmail() string {
	if c == nil {
		return ""
	}
	return c.GDriveEmail
}

// GetGDriveToken returns the Google Drive token from external config
func (c *ExternalTokenConfig) GetGDriveToken() string {
	if c == nil {
		return ""
	}
	return c.GDriveToken
}

// GetGmailToken returns the Gmail token from external config
func (c *ExternalTokenConfig) GetGmailToken() string {
	if c == nil {
		return ""
	}
	return c.GmailToken
}

// GetCalendarToken returns the Calendar token from external config
func (c *ExternalTokenConfig) GetCalendarToken() string {
	if c == nil {
		return ""
	}
	return c.CalendarToken
}

// HasGDriveToken returns true if GDrive token is available
func (c *ExternalTokenConfig) HasGDriveToken() bool {
	return c != nil && c.GDriveToken != ""
}

// HasGmailToken returns true if Gmail token is available
func (c *ExternalTokenConfig) HasGmailToken() bool {
	return c != nil && c.GmailToken != ""
}

// HasCalendarToken returns true if Calendar token is available
func (c *ExternalTokenConfig) HasCalendarToken() bool {
	return c != nil && c.CalendarToken != ""
}
