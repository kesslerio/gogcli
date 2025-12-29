package cmd

import (
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/steipete/gogcli/internal/config"
	"github.com/steipete/gogcli/internal/googleauth"
	"github.com/steipete/gogcli/internal/outfmt"
	"github.com/steipete/gogcli/internal/secrets"
	"github.com/steipete/gogcli/internal/ui"
)

var (
	openSecretsStore = secrets.OpenDefault
	authorizeGoogle  = googleauth.Authorize
)

func newAuthCmd(flags *rootFlags) *cobra.Command {
	cmd := &cobra.Command{
		Use:   "auth",
		Short: "Authentication and accounts",
	}

	cmd.AddCommand(newAuthCredentialsCmd())
	cmd.AddCommand(newAuthAddCmd())
	cmd.AddCommand(newAuthListCmd())
	cmd.AddCommand(newAuthRemoveCmd(flags))
	cmd.AddCommand(newAuthTokensCmd(flags))
	cmd.AddCommand(newAuthManageCmd())
	return cmd
}

func newAuthCredentialsCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "credentials <credentials.json|->",
		Short: "Store OAuth client credentials",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			u := ui.FromContext(cmd.Context())
			inPath := args[0]
			var b []byte
			var err error
			if inPath == "-" {
				b, err = io.ReadAll(os.Stdin)
			} else {
				b, err = os.ReadFile(inPath)
			}
			if err != nil {
				return err
			}

			creds, err := config.ParseGoogleOAuthClientJSON(b)
			if err != nil {
				return err
			}

			if err := config.WriteClientCredentials(creds); err != nil {
				return err
			}

			outPath, _ := config.ClientCredentialsPath()
			if outfmt.IsJSON(cmd.Context()) {
				return outfmt.WriteJSON(os.Stdout, map[string]any{
					"saved": true,
					"path":  outPath,
				})
			}
			u.Out().Printf("path\t%s", outPath)
			return nil
		},
	}
}

func newAuthTokensCmd(flags *rootFlags) *cobra.Command {
	cmd := &cobra.Command{
		Use:   "tokens",
		Short: "Manage stored refresh tokens",
	}

	cmd.AddCommand(&cobra.Command{
		Use:   "list",
		Short: "List stored tokens (by key only)",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			u := ui.FromContext(cmd.Context())
			store, err := openSecretsStore()
			if err != nil {
				return err
			}
			keys, err := store.Keys()
			if err != nil {
				return err
			}
			if len(keys) == 0 {
				if outfmt.IsJSON(cmd.Context()) {
					return outfmt.WriteJSON(os.Stdout, map[string]any{"keys": []string{}})
				}
				u.Err().Println("No tokens stored")
				return nil
			}
			if outfmt.IsJSON(cmd.Context()) {
				return outfmt.WriteJSON(os.Stdout, map[string]any{"keys": keys})
			}
			for _, k := range keys {
				u.Out().Println(k)
			}
			return nil
		},
	})

	cmd.AddCommand(newAuthTokensExportCmd())
	cmd.AddCommand(newAuthTokensImportCmd())

	cmd.AddCommand(&cobra.Command{
		Use:   "delete <email>",
		Short: "Delete a stored refresh token",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			u := ui.FromContext(cmd.Context())
			email := strings.TrimSpace(args[0])
			if email == "" {
				return usage("empty email")
			}

			if err := confirmDestructive(cmd, flags, fmt.Sprintf("delete stored token for %s", email)); err != nil {
				return err
			}

			store, err := openSecretsStore()
			if err != nil {
				return err
			}
			if err := store.DeleteToken(email); err != nil {
				return err
			}
			if outfmt.IsJSON(cmd.Context()) {
				return outfmt.WriteJSON(os.Stdout, map[string]any{
					"deleted": true,
					"email":   email,
				})
			}
			u.Out().Printf("deleted\ttrue")
			u.Out().Printf("email\t%s", email)
			return nil
		},
	})

	return cmd
}

func newAuthTokensExportCmd() *cobra.Command {
	var outPath string
	var overwrite bool

	cmd := &cobra.Command{
		Use:   "export <email>",
		Short: "Export a refresh token to a file (contains secrets)",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			u := ui.FromContext(cmd.Context())
			email := strings.TrimSpace(args[0])
			if email == "" {
				return usage("empty email")
			}
			outPath = strings.TrimSpace(outPath)
			if outPath == "" {
				return usage("empty outPath")
			}

			store, err := openSecretsStore()
			if err != nil {
				return err
			}
			tok, err := store.GetToken(email)
			if err != nil {
				return err
			}

			if mkErr := os.MkdirAll(filepath.Dir(outPath), 0o755); mkErr != nil {
				return mkErr
			}

			flags := os.O_WRONLY | os.O_CREATE | os.O_TRUNC
			if !overwrite {
				flags = os.O_WRONLY | os.O_CREATE | os.O_EXCL
			}
			f, openErr := os.OpenFile(outPath, flags, 0o600)
			if openErr != nil {
				return openErr
			}
			defer func() { _ = f.Close() }()

			type export struct {
				Email        string   `json:"email"`
				Services     []string `json:"services,omitempty"`
				Scopes       []string `json:"scopes,omitempty"`
				CreatedAt    string   `json:"created_at,omitempty"`
				RefreshToken string   `json:"refresh_token"`
			}
			created := ""
			if !tok.CreatedAt.IsZero() {
				created = tok.CreatedAt.UTC().Format(time.RFC3339)
			}

			enc := json.NewEncoder(f)
			enc.SetEscapeHTML(false)
			enc.SetIndent("", "  ")
			if encErr := enc.Encode(export{
				Email:        tok.Email,
				Services:     tok.Services,
				Scopes:       tok.Scopes,
				CreatedAt:    created,
				RefreshToken: tok.RefreshToken,
			}); encErr != nil {
				return encErr
			}

			u.Err().Println("WARNING: exported file contains a refresh token (keep it safe and delete it when done)")
			if outfmt.IsJSON(cmd.Context()) {
				return outfmt.WriteJSON(os.Stdout, map[string]any{
					"exported": true,
					"email":    tok.Email,
					"path":     outPath,
				})
			}
			u.Out().Printf("exported\ttrue")
			u.Out().Printf("email\t%s", tok.Email)
			u.Out().Printf("path\t%s", outPath)
			return nil
		},
	}

	cmd.Flags().StringVar(&outPath, "out", "", "Output file path (required)")
	cmd.Flags().BoolVar(&overwrite, "overwrite", false, "Overwrite output file if it exists")
	return cmd
}

func newAuthTokensImportCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "import <inPath|->",
		Short: "Import a refresh token file into keyring (contains secrets)",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			u := ui.FromContext(cmd.Context())
			inPath := args[0]
			var b []byte
			var err error
			if inPath == "-" {
				b, err = io.ReadAll(os.Stdin)
			} else {
				b, err = os.ReadFile(inPath)
			}
			if err != nil {
				return err
			}

			type export struct {
				Email        string   `json:"email"`
				Services     []string `json:"services,omitempty"`
				Scopes       []string `json:"scopes,omitempty"`
				CreatedAt    string   `json:"created_at,omitempty"`
				RefreshToken string   `json:"refresh_token"`
			}
			var ex export
			if unmarshalErr := json.Unmarshal(b, &ex); unmarshalErr != nil {
				return unmarshalErr
			}
			ex.Email = strings.TrimSpace(ex.Email)
			if ex.Email == "" {
				return usage("missing email in token file")
			}
			if strings.TrimSpace(ex.RefreshToken) == "" {
				return usage("missing refresh_token in token file")
			}
			var createdAt time.Time
			if strings.TrimSpace(ex.CreatedAt) != "" {
				parsed, parseErr := time.Parse(time.RFC3339, strings.TrimSpace(ex.CreatedAt))
				if parseErr != nil {
					return parseErr
				}
				createdAt = parsed
			}

			store, err := openSecretsStore()
			if err != nil {
				return err
			}
			if err := store.SetToken(ex.Email, secrets.Token{
				Email:        ex.Email,
				Services:     ex.Services,
				Scopes:       ex.Scopes,
				CreatedAt:    createdAt,
				RefreshToken: ex.RefreshToken,
			}); err != nil {
				return err
			}

			u.Err().Println("Imported refresh token into keyring")
			if outfmt.IsJSON(cmd.Context()) {
				return outfmt.WriteJSON(os.Stdout, map[string]any{
					"imported": true,
					"email":    ex.Email,
				})
			}
			u.Out().Printf("imported\ttrue")
			u.Out().Printf("email\t%s", ex.Email)
			return nil
		},
	}
	return cmd
}

func newAuthAddCmd() *cobra.Command {
	var manual bool
	var forceConsent bool
	var servicesCSV string

	cmd := &cobra.Command{
		Use:   "add <email>",
		Short: "Authorize and store a refresh token",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			u := ui.FromContext(cmd.Context())

			email := args[0]

			var services []googleauth.Service
			if strings.EqualFold(strings.TrimSpace(servicesCSV), "") || strings.EqualFold(strings.TrimSpace(servicesCSV), "all") {
				services = googleauth.AllServices()
			} else {
				parts := strings.Split(servicesCSV, ",")
				seen := make(map[googleauth.Service]struct{})
				for _, p := range parts {
					svc, err := googleauth.ParseService(p)
					if err != nil {
						return err
					}
					if _, ok := seen[svc]; ok {
						continue
					}
					seen[svc] = struct{}{}
					services = append(services, svc)
				}
			}
			if len(services) == 0 {
				return cmd.Help()
			}

			scopes, err := googleauth.ScopesForServices(services)
			if err != nil {
				return err
			}

			refreshToken, err := authorizeGoogle(cmd.Context(), googleauth.AuthorizeOptions{
				Services:     services,
				Scopes:       scopes,
				Manual:       manual,
				ForceConsent: forceConsent,
			})
			if err != nil {
				return err
			}

			store, err := openSecretsStore()
			if err != nil {
				return err
			}
			serviceNames := make([]string, 0, len(services))
			for _, svc := range services {
				serviceNames = append(serviceNames, string(svc))
			}
			sort.Strings(serviceNames)

			if err := store.SetToken(email, secrets.Token{
				Email:        email,
				Services:     serviceNames,
				Scopes:       scopes,
				RefreshToken: refreshToken,
			}); err != nil {
				return err
			}
			if outfmt.IsJSON(cmd.Context()) {
				return outfmt.WriteJSON(os.Stdout, map[string]any{
					"stored":   true,
					"email":    email,
					"services": serviceNames,
				})
			}
			u.Out().Printf("email\t%s", email)
			u.Out().Printf("services\t%s", strings.Join(serviceNames, ","))
			return nil
		},
	}

	cmd.Flags().BoolVar(&manual, "manual", false, "Browserless auth flow (paste redirect URL)")
	cmd.Flags().BoolVar(&forceConsent, "force-consent", false, "Force consent screen to obtain a refresh token")
	cmd.Flags().StringVar(&servicesCSV, "services", "all", "Services to authorize: all or comma-separated gmail,calendar,drive,contacts,tasks,sheets,people")
	return cmd
}

func newAuthListCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "list",
		Short: "List stored accounts",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			u := ui.FromContext(cmd.Context())
			store, err := openSecretsStore()
			if err != nil {
				return err
			}
			tokens, err := store.ListTokens()
			if err != nil {
				return err
			}
			sort.Slice(tokens, func(i, j int) bool { return tokens[i].Email < tokens[j].Email })
			if outfmt.IsJSON(cmd.Context()) {
				type item struct {
					Email     string   `json:"email"`
					Services  []string `json:"services,omitempty"`
					Scopes    []string `json:"scopes,omitempty"`
					CreatedAt string   `json:"created_at,omitempty"`
				}
				out := make([]item, 0, len(tokens))
				for _, t := range tokens {
					created := ""
					if !t.CreatedAt.IsZero() {
						created = t.CreatedAt.UTC().Format("2006-01-02T15:04:05Z07:00")
					}
					out = append(out, item{
						Email:     t.Email,
						Services:  t.Services,
						Scopes:    t.Scopes,
						CreatedAt: created,
					})
				}
				return outfmt.WriteJSON(os.Stdout, map[string]any{"accounts": out})
			}
			if len(tokens) == 0 {
				u.Err().Println("No tokens stored")
				return nil
			}
			for _, t := range tokens {
				created := ""
				if !t.CreatedAt.IsZero() {
					created = t.CreatedAt.UTC().Format("2006-01-02T15:04:05Z07:00")
				}
				u.Out().Printf("%s\t%s\t%s", t.Email, strings.Join(t.Services, ","), created)
			}
			return nil
		},
	}
}

func newAuthRemoveCmd(flags *rootFlags) *cobra.Command {
	return &cobra.Command{
		Use:   "remove <email>",
		Short: "Remove a stored refresh token",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			u := ui.FromContext(cmd.Context())
			email := strings.TrimSpace(args[0])
			if email == "" {
				return usage("empty email")
			}

			if err := confirmDestructive(cmd, flags, fmt.Sprintf("remove stored token for %s", email)); err != nil {
				return err
			}
			store, err := openSecretsStore()
			if err != nil {
				return err
			}
			if err := store.DeleteToken(email); err != nil {
				return err
			}
			if outfmt.IsJSON(cmd.Context()) {
				return outfmt.WriteJSON(os.Stdout, map[string]any{
					"deleted": true,
					"email":   email,
				})
			}
			u.Out().Printf("deleted\ttrue")
			u.Out().Printf("email\t%s", email)
			return nil
		},
	}
}

func newAuthManageCmd() *cobra.Command {
	var forceConsent bool
	var servicesCSV string
	var timeout time.Duration

	cmd := &cobra.Command{
		Use:     "manage",
		Aliases: []string{"login"},
		Short:   "Open accounts manager in browser",
		Long:    "Opens a browser-based UI to manage Google accounts, add new accounts, set defaults, and remove accounts.",
		Args:    cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			var services []googleauth.Service
			if strings.EqualFold(strings.TrimSpace(servicesCSV), "") || strings.EqualFold(strings.TrimSpace(servicesCSV), "all") {
				services = googleauth.AllServices()
			} else {
				parts := strings.Split(servicesCSV, ",")
				seen := make(map[googleauth.Service]struct{})
				for _, p := range parts {
					svc, err := googleauth.ParseService(p)
					if err != nil {
						return err
					}
					if _, ok := seen[svc]; ok {
						continue
					}
					seen[svc] = struct{}{}
					services = append(services, svc)
				}
			}

			return googleauth.StartManageServer(cmd.Context(), googleauth.ManageServerOptions{
				Timeout:      timeout,
				Services:     services,
				ForceConsent: forceConsent,
			})
		},
	}

	cmd.Flags().BoolVar(&forceConsent, "force-consent", false, "Force consent screen when adding accounts")
	cmd.Flags().StringVar(&servicesCSV, "services", "all", "Services to authorize: all or comma-separated gmail,calendar,drive,contacts,tasks,sheets,people")
	cmd.Flags().DurationVar(&timeout, "timeout", 10*time.Minute, "Server timeout duration")
	return cmd
}
