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
	GDriveToken string // token for Google Drive
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
			}
		}
	}

	return config, nil
}

// GetGDriveToken returns the Google Drive token from external config
func (c *ExternalTokenConfig) GetGDriveToken() string {
	if c == nil {
		return ""
	}
	return c.GDriveToken
}

// HasGDriveToken returns true if GDrive token is available
func (c *ExternalTokenConfig) HasGDriveToken() bool {
	return c != nil && c.GDriveToken != ""
}
