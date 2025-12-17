package cmd

import (
	"encoding/base64"
	"errors"
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/steipete/gogcli/internal/outfmt"
	"github.com/steipete/gogcli/internal/ui"
)

func newGmailGetCmd(flags *rootFlags) *cobra.Command {
	var format string
	var headers string

	cmd := &cobra.Command{
		Use:   "get <messageId>",
		Short: "Get a message (full|metadata|raw)",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			u := ui.FromContext(cmd.Context())
			account, err := requireAccount(flags)
			if err != nil {
				return err
			}
			messageID := strings.TrimSpace(args[0])
			if messageID == "" {
				return errors.New("empty messageId")
			}

			format = strings.TrimSpace(format)
			if format == "" {
				format = "full"
			}
			switch format {
			case "full", "metadata", "raw":
			default:
				return fmt.Errorf("invalid --format: %q (expected full|metadata|raw)", format)
			}

			svc, err := newGmailService(cmd.Context(), account)
			if err != nil {
				return err
			}

			call := svc.Users.Messages.Get("me", messageID).Format(format).Context(cmd.Context())
			if format == "metadata" {
				headerList := splitCSV(headers)
				if len(headerList) == 0 {
					headerList = []string{"From", "To", "Subject", "Date"}
				}
				call = call.MetadataHeaders(headerList...)
			}

			msg, err := call.Do()
			if err != nil {
				return err
			}

			if outfmt.IsJSON(cmd.Context()) {
				return outfmt.WriteJSON(os.Stdout, map[string]any{"message": msg})
			}

			u.Out().Printf("id\t%s", msg.Id)
			u.Out().Printf("thread_id\t%s", msg.ThreadId)
			u.Out().Printf("label_ids\t%s", strings.Join(msg.LabelIds, ","))

			switch format {
			case "raw":
				if msg.Raw == "" {
					u.Err().Println("Empty raw message")
					return nil
				}
				decoded, err := base64.RawURLEncoding.DecodeString(msg.Raw)
				if err != nil {
					return err
				}
				u.Out().Println("")
				u.Out().Println(string(decoded))
				return nil
			case "metadata", "full":
				u.Out().Printf("from\t%s", headerValue(msg.Payload, "From"))
				u.Out().Printf("to\t%s", headerValue(msg.Payload, "To"))
				u.Out().Printf("subject\t%s", headerValue(msg.Payload, "Subject"))
				u.Out().Printf("date\t%s", headerValue(msg.Payload, "Date"))
				if format == "full" {
					body := bestBodyText(msg.Payload)
					if body != "" {
						u.Out().Println("")
						u.Out().Println(body)
					}
				}
				return nil
			default:
				return nil
			}
		},
	}

	cmd.Flags().StringVar(&format, "format", "full", "Message format: full|metadata|raw")
	cmd.Flags().StringVar(&headers, "headers", "", "Metadata headers (comma-separated; only for --format=metadata)")
	return cmd
}
