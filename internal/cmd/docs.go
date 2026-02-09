package cmd

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"strings"

	"google.golang.org/api/docs/v1"
	"google.golang.org/api/drive/v3"
	gapi "google.golang.org/api/googleapi"

	"github.com/steipete/gogcli/internal/googleapi"
	"github.com/steipete/gogcli/internal/outfmt"
	"github.com/steipete/gogcli/internal/ui"
)

var newDocsService = googleapi.NewDocs

type DocsCmd struct {
	Export DocsExportCmd `cmd:"" name:"export" help:"Export a Google Doc (pdf|docx|txt)"`
	Info   DocsInfoCmd   `cmd:"" name:"info" help:"Get Google Doc metadata"`
	Create DocsCreateCmd `cmd:"" name:"create" help:"Create a Google Doc"`
	Copy   DocsCopyCmd   `cmd:"" name:"copy" help:"Copy a Google Doc"`
	Cat    DocsCatCmd    `cmd:"" name:"cat" help:"Print a Google Doc as plain text"`
	Write  DocsWriteCmd  `cmd:"" name:"write" help:"Write markdown content to a Google Doc"`
	Append DocsAppendCmd `cmd:"" name:"append" help:"Append markdown content to a Google Doc"`
	Clear  DocsClearCmd  `cmd:"" name:"clear" help:"Clear all content from a Google Doc"`
}

type DocsExportCmd struct {
	DocID  string         `arg:"" name:"docId" help:"Doc ID"`
	Output OutputPathFlag `embed:""`
	Format string         `name:"format" help:"Export format: pdf|docx|txt" default:"pdf"`
}

func (c *DocsExportCmd) Run(ctx context.Context, flags *RootFlags) error {
	return exportViaDrive(ctx, flags, exportViaDriveOptions{
		ArgName:       "docId",
		ExpectedMime:  "application/vnd.google-apps.document",
		KindLabel:     "Google Doc",
		DefaultFormat: "pdf",
	}, c.DocID, c.Output.Path, c.Format)
}

type DocsInfoCmd struct {
	DocID string `arg:"" name:"docId" help:"Doc ID"`
}

func (c *DocsInfoCmd) Run(ctx context.Context, flags *RootFlags) error {
	u := ui.FromContext(ctx)
	account, err := requireAccount(flags)
	if err != nil {
		return err
	}

	id := strings.TrimSpace(c.DocID)
	if id == "" {
		return usage("empty docId")
	}

	svc, err := newDocsService(ctx, account)
	if err != nil {
		return err
	}

	doc, err := svc.Documents.Get(id).
		Fields("documentId,title,revisionId").
		Context(ctx).
		Do()
	if err != nil {
		if isDocsNotFound(err) {
			return fmt.Errorf("doc not found or not a Google Doc (id=%s)", id)
		}
		return err
	}
	if doc == nil {
		return errors.New("doc not found")
	}

	file := map[string]any{
		"id":       doc.DocumentId,
		"name":     doc.Title,
		"mimeType": driveMimeGoogleDoc,
	}
	if link := docsWebViewLink(doc.DocumentId); link != "" {
		file["webViewLink"] = link
	}

	if outfmt.IsJSON(ctx) {
		return outfmt.WriteJSON(os.Stdout, map[string]any{
			strFile:    file,
			"document": doc,
		})
	}

	u.Out().Printf("id\t%s", doc.DocumentId)
	u.Out().Printf("name\t%s", doc.Title)
	u.Out().Printf("mime\t%s", driveMimeGoogleDoc)
	if link := docsWebViewLink(doc.DocumentId); link != "" {
		u.Out().Printf("link\t%s", link)
	}
	if doc.RevisionId != "" {
		u.Out().Printf("revision\t%s", doc.RevisionId)
	}
	return nil
}

type DocsCreateCmd struct {
	Title  string `arg:"" name:"title" help:"Doc title"`
	Parent string `name:"parent" help:"Destination folder ID"`
}

func (c *DocsCreateCmd) Run(ctx context.Context, flags *RootFlags) error {
	u := ui.FromContext(ctx)
	account, err := requireAccount(flags)
	if err != nil {
		return err
	}

	title := strings.TrimSpace(c.Title)
	if title == "" {
		return usage("empty title")
	}

	svc, err := newDriveService(ctx, account)
	if err != nil {
		return err
	}

	f := &drive.File{
		Name:     title,
		MimeType: "application/vnd.google-apps.document",
	}
	parent := strings.TrimSpace(c.Parent)
	if parent != "" {
		f.Parents = []string{parent}
	}

	created, err := svc.Files.Create(f).
		SupportsAllDrives(true).
		Fields("id, name, mimeType, webViewLink").
		Context(ctx).
		Do()
	if err != nil {
		return err
	}
	if created == nil {
		return errors.New("create failed")
	}

	if outfmt.IsJSON(ctx) {
		return outfmt.WriteJSON(os.Stdout, map[string]any{strFile: created})
	}

	u.Out().Printf("id\t%s", created.Id)
	u.Out().Printf("name\t%s", created.Name)
	u.Out().Printf("mime\t%s", created.MimeType)
	if created.WebViewLink != "" {
		u.Out().Printf("link\t%s", created.WebViewLink)
	}
	return nil
}

type DocsCopyCmd struct {
	DocID  string `arg:"" name:"docId" help:"Doc ID"`
	Title  string `arg:"" name:"title" help:"New title"`
	Parent string `name:"parent" help:"Destination folder ID"`
}

func (c *DocsCopyCmd) Run(ctx context.Context, flags *RootFlags) error {
	return copyViaDrive(ctx, flags, copyViaDriveOptions{
		ArgName:      "docId",
		ExpectedMime: "application/vnd.google-apps.document",
		KindLabel:    "Google Doc",
	}, c.DocID, c.Title, c.Parent)
}

type DocsCatCmd struct {
	DocID    string `arg:"" name:"docId" help:"Doc ID"`
	MaxBytes int64  `name:"max-bytes" help:"Max bytes to read (0 = unlimited)" default:"2000000"`
}

func (c *DocsCatCmd) Run(ctx context.Context, flags *RootFlags) error {
	account, err := requireAccount(flags)
	if err != nil {
		return err
	}

	id := strings.TrimSpace(c.DocID)
	if id == "" {
		return usage("empty docId")
	}

	svc, err := newDocsService(ctx, account)
	if err != nil {
		return err
	}

	doc, err := svc.Documents.Get(id).
		Context(ctx).
		Do()
	if err != nil {
		if isDocsNotFound(err) {
			return fmt.Errorf("doc not found or not a Google Doc (id=%s)", id)
		}
		return err
	}
	if doc == nil {
		return errors.New("doc not found")
	}

	text := docsPlainText(doc, c.MaxBytes)

	if outfmt.IsJSON(ctx) {
		return outfmt.WriteJSON(os.Stdout, map[string]any{"text": text})
	}
	_, err = io.WriteString(os.Stdout, text)
	return err
}

func docsWebViewLink(id string) string {
	id = strings.TrimSpace(id)
	if id == "" {
		return ""
	}
	return "https://docs.google.com/document/d/" + id + "/edit"
}

func docsPlainText(doc *docs.Document, maxBytes int64) string {
	if doc == nil || doc.Body == nil {
		return ""
	}

	var buf bytes.Buffer
	for _, el := range doc.Body.Content {
		if !appendDocsElementText(&buf, maxBytes, el) {
			break
		}
	}

	return buf.String()
}

func appendDocsElementText(buf *bytes.Buffer, maxBytes int64, el *docs.StructuralElement) bool {
	if el == nil {
		return true
	}

	switch {
	case el.Paragraph != nil:
		for _, p := range el.Paragraph.Elements {
			if p.TextRun == nil {
				continue
			}
			if !appendLimited(buf, maxBytes, p.TextRun.Content) {
				return false
			}
		}
	case el.Table != nil:
		for rowIdx, row := range el.Table.TableRows {
			if rowIdx > 0 {
				if !appendLimited(buf, maxBytes, "\n") {
					return false
				}
			}
			for cellIdx, cell := range row.TableCells {
				if cellIdx > 0 {
					if !appendLimited(buf, maxBytes, "\t") {
						return false
					}
				}
				for _, content := range cell.Content {
					if !appendDocsElementText(buf, maxBytes, content) {
						return false
					}
				}
			}
		}
	case el.TableOfContents != nil:
		for _, content := range el.TableOfContents.Content {
			if !appendDocsElementText(buf, maxBytes, content) {
				return false
			}
		}
	}

	return true
}

func appendLimited(buf *bytes.Buffer, maxBytes int64, s string) bool {
	if maxBytes <= 0 {
		_, _ = buf.WriteString(s)
		return true
	}

	remaining := int(maxBytes) - buf.Len()
	if remaining <= 0 {
		return false
	}
	if len(s) > remaining {
		_, _ = buf.WriteString(s[:remaining])
		return false
	}
	_, _ = buf.WriteString(s)
	return true
}

func isDocsNotFound(err error) bool {
	var apiErr *gapi.Error
	if !errors.As(err, &apiErr) {
		return false
	}
	return apiErr.Code == http.StatusNotFound
}

// utf16CodeUnitCount returns the number of UTF-16 code units in a string.
// Google Docs API uses UTF-16 code unit offsets for formatting ranges.
// Characters outside the Basic Multilingual Plane (like emojis) are encoded
// as surrogate pairs and count as 2 code units.
func utf16CodeUnitCount(s string) int64 {
	var count int64
	for _, r := range s {
		if r >= 0x10000 {
			count += 2 // Surrogate pair
		} else {
			count++
		}
	}
	return count
}

// DocsWriteCmd writes markdown content to a Google Doc.
type DocsWriteCmd struct {
	DocID string `arg:"" name:"docId" help:"Doc ID"`
	File  string `name:"file" help:"Markdown file to write (or stdin if omitted)" type:"existingfile"`
}

func (c *DocsWriteCmd) Run(ctx context.Context, flags *RootFlags) error {
	u := ui.FromContext(ctx)
	account, err := requireAccount(flags)
	if err != nil {
		return err
	}

	id := strings.TrimSpace(c.DocID)
	if id == "" {
		return usage("empty docId")
	}

	svc, err := newDocsService(ctx, account)
	if err != nil {
		return err
	}

	// Read input from file or stdin
	input, err := readInput(c.File)
	if err != nil {
		return err
	}

	// Clear the document first
	if err := clearDocsContent(ctx, svc, id); err != nil {
		return fmt.Errorf("clear doc content: %w", err)
	}

	// Parse markdown and apply to doc
	if err := writeMarkdownToDoc(ctx, svc, id, string(input), 1); err != nil {
		return err
	}

	if outfmt.IsJSON(ctx) {
		return outfmt.WriteJSON(os.Stdout, map[string]any{
			"documentId": id,
		})
	}

	u.Out().Printf("wrote document %s", id)
	return nil
}

// DocsAppendCmd appends markdown content to a Google Doc.
type DocsAppendCmd struct {
	DocID string `arg:"" name:"docId" help:"Doc ID"`
	File  string `name:"file" help:"Markdown file to append (or stdin if omitted)" type:"existingfile"`
}

func (c *DocsAppendCmd) Run(ctx context.Context, flags *RootFlags) error {
	u := ui.FromContext(ctx)
	account, err := requireAccount(flags)
	if err != nil {
		return err
	}

	id := strings.TrimSpace(c.DocID)
	if id == "" {
		return usage("empty docId")
	}

	svc, err := newDocsService(ctx, account)
	if err != nil {
		return err
	}

	// Read input from file or stdin
	input, err := readInput(c.File)
	if err != nil {
		return err
	}

	// Get current document to find end index
	doc, err := svc.Documents.Get(id).Context(ctx).Do()
	if err != nil {
		if isDocsNotFound(err) {
			return fmt.Errorf("doc not found or not a Google Doc (id=%s)", id)
		}
		return err
	}

	endIndex := int64(1)
	if doc.Body != nil && len(doc.Body.Content) > 0 {
		endIndex = doc.Body.Content[len(doc.Body.Content)-1].EndIndex
	}

	// Parse markdown and apply to doc
	if err := writeMarkdownToDoc(ctx, svc, id, string(input), endIndex); err != nil {
		return err
	}

	if outfmt.IsJSON(ctx) {
		return outfmt.WriteJSON(os.Stdout, map[string]any{
			"documentId": id,
		})
	}

	u.Out().Printf("appended to document %s", id)
	return nil
}

// DocsClearCmd clears all content from a document.
type DocsClearCmd struct {
	DocID string `arg:"" name:"docId" help:"Doc ID"`
}

func (c *DocsClearCmd) Run(ctx context.Context, flags *RootFlags) error {
	u := ui.FromContext(ctx)
	account, err := requireAccount(flags)
	if err != nil {
		return err
	}

	id := strings.TrimSpace(c.DocID)
	if id == "" {
		return usage("empty docId")
	}

	svc, err := newDocsService(ctx, account)
	if err != nil {
		return err
	}

	if err := clearDocsContent(ctx, svc, id); err != nil {
		return err
	}

	if outfmt.IsJSON(ctx) {
		return outfmt.WriteJSON(os.Stdout, map[string]any{
			"documentId": id,
		})
	}

	u.Out().Printf("cleared document %s", id)
	return nil
}

func readInput(file string) ([]byte, error) {
	if strings.TrimSpace(file) != "" {
		return os.ReadFile(file)
	}
	return io.ReadAll(os.Stdin)
}

// clearDocsContent deletes all content from a document.
func clearDocsContent(ctx context.Context, svc *docs.Service, docID string) error {
	doc, err := svc.Documents.Get(docID).Context(ctx).Do()
	if err != nil {
		return err
	}

	if doc.Body == nil || len(doc.Body.Content) == 0 {
		return nil
	}

	// Get the end index of the document
	var endIndex int64
	for _, el := range doc.Body.Content {
		if el.EndIndex > endIndex {
			endIndex = el.EndIndex
		}
	}

	if endIndex <= 2 {
		return nil // Document is already empty (just has the mandatory final newline)
	}

	// Delete from index 1 to endIndex - 1.
	// Google Docs always has a final newline at the end with index endIndex.
	// We delete everything up to (but not including) that final newline.
	req := &docs.BatchUpdateDocumentRequest{
		Requests: []*docs.Request{
			{
				DeleteContentRange: &docs.DeleteContentRangeRequest{
					Range: &docs.Range{
						StartIndex: 1,
						EndIndex:   endIndex - 1,
					},
				},
			},
		},
	}

	_, err = svc.Documents.BatchUpdate(docID, req).Context(ctx).Do()
	return err
}

// writeMarkdownToDoc parses markdown and writes it to a Google Doc.
func writeMarkdownToDoc(ctx context.Context, svc *docs.Service, docID string, markdown string, startIdx int64) error {
	type formatRange struct {
		start int64
		end   int64
		bold  bool
		italic bool
	}
	type segment struct {
		text         string
		style        string // HEADING_1, HEADING_2, HEADING_3, NORMAL_TEXT
		isBullet     bool
		isNumbered   bool
		ranges       []formatRange
	}

	lines := strings.Split(markdown, "\n")
	var segments []segment

	for _, line := range lines {
		line = strings.TrimRight(line, "\r")
		trimmed := strings.TrimSpace(line)

		if trimmed == "" {
			segments = append(segments, segment{text: "\n", style: "NORMAL_TEXT"})
			continue
		}

		if trimmed == "---" {
			segments = append(segments, segment{text: "\n", style: "NORMAL_TEXT"})
			continue
		}

		var s segment
		text := line

		// Check headings
		switch {
		case strings.HasPrefix(trimmed, "### "):
			s.style = "HEADING_3"
			text = strings.TrimPrefix(line, "### ")
		case strings.HasPrefix(trimmed, "## "):
			s.style = "HEADING_2"
			text = strings.TrimPrefix(line, "## ")
		case strings.HasPrefix(trimmed, "# "):
			s.style = "HEADING_1"
			text = strings.TrimPrefix(line, "# ")
		default:
			s.style = "NORMAL_TEXT"
		}

		// Check list items (only for NORMAL_TEXT, not headings)
		if s.style == "NORMAL_TEXT" {
			if strings.HasPrefix(text, "- ") || strings.HasPrefix(text, "* ") {
				s.isBullet = true
				text = text[2:]
			} else if len(text) >= 3 && text[0] >= '0' && text[0] <= '9' && text[1] == '.' && text[2] == ' ' {
				s.isNumbered = true
				text = text[3:]
			}
		}

		// Simple bold/italic parser with UTF-16 code unit awareness
		// Google Docs API expects UTF-16 code unit offsets, not byte offsets
		var cleanText strings.Builder
		var currentOffset int64
		
		tempText := text
		for {
			boldStart := strings.Index(tempText, "**")
			italicStart := strings.Index(tempText, "*")
			
			if boldStart == -1 && italicStart == -1 {
				cleanText.WriteString(tempText)
				break
			}
			
			if boldStart != -1 && (italicStart == -1 || boldStart <= italicStart) {
				// Process text before the bold marker
				beforeBold := tempText[:boldStart]
				cleanText.WriteString(beforeBold)
				currentOffset += utf16CodeUnitCount(beforeBold)
				
				tempText = tempText[boldStart+2:]
				boldEnd := strings.Index(tempText, "**")
				if boldEnd != -1 {
					innerText := tempText[:boldEnd]
					s.ranges = append(s.ranges, formatRange{
						start: currentOffset,
						end:   currentOffset + utf16CodeUnitCount(innerText),
						bold:  true,
					})
					cleanText.WriteString(innerText)
					currentOffset += utf16CodeUnitCount(innerText)
					tempText = tempText[boldEnd+2:]
				} else {
					cleanText.WriteString("**")
					currentOffset += 2
				}
			} else if italicStart != -1 {
				// Process text before the italic marker
				beforeItalic := tempText[:italicStart]
				cleanText.WriteString(beforeItalic)
				currentOffset += utf16CodeUnitCount(beforeItalic)
				
				tempText = tempText[italicStart+1:]
				italicEnd := strings.Index(tempText, "*")
				if italicEnd != -1 {
					innerText := tempText[:italicEnd]
					s.ranges = append(s.ranges, formatRange{
						start: currentOffset,
						end:   currentOffset + utf16CodeUnitCount(innerText),
						italic: true,
					})
					cleanText.WriteString(innerText)
					currentOffset += utf16CodeUnitCount(innerText)
					tempText = tempText[italicEnd+1:]
				} else {
					cleanText.WriteString("*")
					currentOffset += 1
				}
			}
		}

		s.text = cleanText.String() + "\n"
		segments = append(segments, s)
	}

	// Build all text and insert at once
	var fullText strings.Builder
	for _, seg := range segments {
		fullText.WriteString(seg.text)
	}

	// Insert all text at the starting index
	req := &docs.BatchUpdateDocumentRequest{
		Requests: []*docs.Request{
			{
				InsertText: &docs.InsertTextRequest{
					Location: &docs.Location{Index: startIdx},
					Text:     fullText.String(),
				},
			},
		},
	}

	_, err := svc.Documents.BatchUpdate(docID, req).Context(ctx).Do()
	if err != nil {
		return fmt.Errorf("insert text: %w", err)
	}

	// Now apply styles
	var styleRequests []*docs.Request
	idx := startIdx

	for _, seg := range segments {
		segLen := int64(len(seg.text))
		paraStartIdx := idx
		paraEndIdx := idx + segLen

		if seg.style != "" && seg.style != "NORMAL_TEXT" {
			styleRequests = append(styleRequests, &docs.Request{
				UpdateParagraphStyle: &docs.UpdateParagraphStyleRequest{
					Range: &docs.Range{
						StartIndex: paraStartIdx,
						EndIndex:   paraEndIdx,
					},
					ParagraphStyle: &docs.ParagraphStyle{
						NamedStyleType: seg.style,
					},
					Fields: "namedStyleType",
				},
			})
		}

		if seg.isBullet {
			styleRequests = append(styleRequests, &docs.Request{
				CreateParagraphBullets: &docs.CreateParagraphBulletsRequest{
					Range: &docs.Range{
						StartIndex: paraStartIdx,
						EndIndex:   paraEndIdx,
					},
					BulletPreset: "BULLET_DISC_CIRCLE_SQUARE",
				},
			})
		} else if seg.isNumbered {
			styleRequests = append(styleRequests, &docs.Request{
				CreateParagraphBullets: &docs.CreateParagraphBulletsRequest{
					Range: &docs.Range{
						StartIndex: paraStartIdx,
						EndIndex:   paraEndIdx,
					},
					BulletPreset: "NUMBERED_DECIMAL_NESTED",
				},
			})
		}

		for _, r := range seg.ranges {
			textStyle := &docs.TextStyle{}
			var fields []string
			if r.bold {
				textStyle.Bold = true
				fields = append(fields, "bold")
			}
			if r.italic {
				textStyle.Italic = true
				fields = append(fields, "italic")
			}
			styleRequests = append(styleRequests, &docs.Request{
				UpdateTextStyle: &docs.UpdateTextStyleRequest{
					Range: &docs.Range{
						StartIndex: paraStartIdx + r.start,
						EndIndex:   paraStartIdx + r.end,
					},
					TextStyle: textStyle,
					Fields:    strings.Join(fields, ","),
				},
			})
		}

		idx = paraEndIdx
	}

	if len(styleRequests) > 0 {
		// Google Docs API has a limit of 50 requests per batchUpdate call.
		const batchSize = 50
		for i := 0; i < len(styleRequests); i += batchSize {
			end := i + batchSize
			if end > len(styleRequests) {
				end = len(styleRequests)
			}
			batchReq := &docs.BatchUpdateDocumentRequest{
				Requests: styleRequests[i:end],
			}
			_, err = svc.Documents.BatchUpdate(docID, batchReq).Context(ctx).Do()
			if err != nil {
				return fmt.Errorf("apply styles batch: %w", err)
			}
		}
	}

	return nil
}
