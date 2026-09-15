package cmd

import (
	"context"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"testing"

	"google.golang.org/api/calendar/v3"
	"google.golang.org/api/option"
)

func TestCalendarEventsListCall_DefaultsToHideCancelledEvents(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.URL.Query().Get("showDeleted"); got != "false" {
			t.Fatalf("expected showDeleted=false, got %q", got)
		}
		if got := r.URL.Query().Get("singleEvents"); got != "true" {
			t.Fatalf("expected singleEvents=true, got %q", got)
		}

		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]any{"items": []any{}})
	}))
	defer srv.Close()

	svc, err := calendar.NewService(context.Background(),
		option.WithHTTPClient(srv.Client()),
		option.WithEndpoint(srv.URL+"/"),
		option.WithoutAuthentication(),
	)
	if err != nil {
		t.Fatalf("NewService: %v", err)
	}

	if _, err := calendarEventsListCall(context.Background(), svc, "primary", "2026-01-01T00:00:00Z", "2026-01-02T00:00:00Z", 10, "", "", "", "", "", false).Do(); err != nil {
		t.Fatalf("Do: %v", err)
	}
}

func TestCalendarEventsListCall_ShowsCancelledEventsWhenRequested(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.URL.Query().Get("showDeleted"); got != "true" {
			t.Fatalf("expected showDeleted=true, got %q", got)
		}
		if got := r.URL.Query().Get("singleEvents"); got != "true" {
			t.Fatalf("expected singleEvents=true, got %q", got)
		}

		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]any{"items": []any{}})
	}))
	defer srv.Close()

	svc, err := calendar.NewService(context.Background(),
		option.WithHTTPClient(srv.Client()),
		option.WithEndpoint(srv.URL+"/"),
		option.WithoutAuthentication(),
	)
	if err != nil {
		t.Fatalf("NewService: %v", err)
	}

	if _, err := calendarEventsListCall(context.Background(), svc, "primary", "2026-01-01T00:00:00Z", "2026-01-02T00:00:00Z", 10, "", "", "", "", "", true).Do(); err != nil {
		t.Fatalf("Do: %v", err)
	}
}

func TestListCalendarEvents_ShowDeletedAllPagesPreservesCancellationDetails(t *testing.T) {
	pageCalls := 0
	svc, closeServer := newCalendarServiceForTest(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.URL.Query().Get("showDeleted"); got != "true" {
			t.Fatalf("expected showDeleted=true, got %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		switch r.URL.Query().Get("pageToken") {
		case "":
			pageCalls++
			_ = json.NewEncoder(w).Encode(map[string]any{
				"items": []map[string]any{{
					"id":          "deleted-event",
					"status":      "cancelled",
					"updated":     "2026-01-01T12:00:00Z",
					"summary":     "Cancelled consultation",
					"description": "Customer requested cancellation",
					"location":    "Video call",
					"attendees":   []map[string]any{{"email": "buyer@example.com"}},
				}},
				"nextPageToken": "page-2",
			})
		case "page-2":
			pageCalls++
			_ = json.NewEncoder(w).Encode(map[string]any{
				"items": []map[string]any{{
					"id":               "cancelled-instance",
					"status":           "cancelled",
					"recurringEventId": "recurring-series",
					"originalStartTime": map[string]any{
						"dateTime": "2026-01-02T10:00:00Z",
					},
					"updated": "2026-01-01T13:00:00Z",
				}},
			})
		default:
			t.Fatalf("unexpected page token %q", r.URL.Query().Get("pageToken"))
		}
	}))
	defer closeServer()

	ctx := newCalendarJSONContext(t)
	jsonOut := captureStdout(t, func() {
		if err := listCalendarEvents(ctx, svc, "cal1", "2026-01-01T00:00:00Z", "2026-01-03T00:00:00Z", 1, "", true, false, "", "", "", "", false, true); err != nil {
			t.Fatalf("listCalendarEvents: %v", err)
		}
	})

	var parsed struct {
		Events []map[string]any `json:"events"`
	}
	if err := json.Unmarshal([]byte(jsonOut), &parsed); err != nil {
		t.Fatalf("json parse: %v", err)
	}
	if pageCalls != 2 || len(parsed.Events) != 2 {
		t.Fatalf("expected two exhausted pages and events, calls=%d events=%#v", pageCalls, parsed.Events)
	}

	deleted := parsed.Events[0]
	for key, want := range map[string]any{
		"id":          "deleted-event",
		"status":      "cancelled",
		"updated":     "2026-01-01T12:00:00Z",
		"summary":     "Cancelled consultation",
		"description": "Customer requested cancellation",
		"location":    "Video call",
	} {
		if got := deleted[key]; got != want {
			t.Errorf("deleted event %s=%#v, want %#v", key, got, want)
		}
	}
	attendees, ok := deleted["attendees"].([]any)
	if !ok || len(attendees) != 1 {
		t.Fatalf("deleted event attendees=%#v", deleted["attendees"])
	}
	attendee, ok := attendees[0].(map[string]any)
	if !ok || attendee["email"] != "buyer@example.com" {
		t.Fatalf("deleted event attendee=%#v", attendees[0])
	}

	instance := parsed.Events[1]
	for key, want := range map[string]any{
		"id":               "cancelled-instance",
		"status":           "cancelled",
		"recurringEventId": "recurring-series",
		"updated":          "2026-01-01T13:00:00Z",
	} {
		if got := instance[key]; got != want {
			t.Errorf("cancelled instance %s=%#v, want %#v", key, got, want)
		}
	}
	originalStart, ok := instance["originalStartTime"].(map[string]any)
	if !ok || originalStart["dateTime"] != "2026-01-02T10:00:00Z" {
		t.Fatalf("cancelled instance originalStartTime=%#v", instance["originalStartTime"])
	}
}
