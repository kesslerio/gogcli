package cmd

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"net/http/httptest"
	"slices"
	"strconv"
	"strings"
	"testing"

	"google.golang.org/api/calendar/v3"
	"google.golang.org/api/option"
)

func TestCalendarEventsListCall_ShowDeletedOption(t *testing.T) {
	for _, showDeleted := range []bool{false, true} {
		t.Run(fmt.Sprintf("show_deleted_%t", showDeleted), func(t *testing.T) {
			srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
				if got := r.URL.Query().Get("showDeleted"); got != strconv.FormatBool(showDeleted) {
					t.Fatalf("expected showDeleted=%t, got %q", showDeleted, got)
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
			if _, err := calendarEventsListCall(context.Background(), svc, "primary", "2026-01-01T00:00:00Z", "2026-01-02T00:00:00Z", 10, "", "", "", "", nil, "", showDeleted).Do(); err != nil {
				t.Fatalf("Do: %v", err)
			}
		})
	}
}

func TestListCalendarEvents_ShowDeletedAllPagesPreservesCancellationDetails(t *testing.T) {
	pageCalls := 0
	svc, closeServer := newCalendarServiceForTest(t, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if !strings.HasSuffix(r.URL.Path, "/events") {
			w.Header().Set("Content-Type", "application/json")
			_ = json.NewEncoder(w).Encode(map[string]any{"id": "cal1", "timeZone": "UTC"})
			return
		}
		if got := r.URL.Query().Get("showDeleted"); got != "true" {
			t.Fatalf("expected showDeleted=true, got %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		switch r.URL.Query().Get("pageToken") {
		case "":
			pageCalls++
			_ = json.NewEncoder(w).Encode(map[string]any{
				"items": []map[string]any{{
					"id": "deleted-event", "status": "cancelled",
					"updated": "2026-01-01T12:00:00Z", "summary": "Cancelled consultation",
					"description": "Customer requested cancellation", "location": "Video call",
					"attendees": []map[string]any{{"email": "buyer@example.com"}},
				}},
				"nextPageToken": "page-2",
			})
		case "page-2":
			pageCalls++
			_ = json.NewEncoder(w).Encode(map[string]any{
				"items": []map[string]any{{
					"id": "cancelled-instance", "status": "cancelled",
					"recurringEventId":  "recurring-series",
					"originalStartTime": map[string]any{"dateTime": "2026-01-02T10:00:00Z"},
					"updated":           "2026-01-01T13:00:00Z",
				}},
			})
		default:
			t.Fatalf("unexpected page token %q", r.URL.Query().Get("pageToken"))
		}
	}))
	defer closeServer()

	ctx := newCmdJSONContext(t)
	jsonOut := captureStdout(t, func() {
		if err := listCalendarEvents(ctx, svc, "cal1", "2026-01-01T00:00:00Z", "2026-01-03T00:00:00Z", 1, "", true, false, "", "", "", "", nil, false, false, "", "", true); err != nil {
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
		"id": "deleted-event", "status": "cancelled", "updated": "2026-01-01T12:00:00Z",
		"summary": "Cancelled consultation", "description": "Customer requested cancellation", "location": "Video call",
	} {
		if got := deleted[key]; got != want {
			t.Errorf("deleted event %s=%#v, want %#v", key, got, want)
		}
	}
	instance := parsed.Events[1]
	if instance["id"] != "cancelled-instance" || instance["status"] != "cancelled" ||
		instance["recurringEventId"] != "recurring-series" || instance["updated"] != "2026-01-01T13:00:00Z" {
		t.Fatalf("cancelled instance fields=%#v", instance)
	}
	originalStart, ok := instance["originalStartTime"].(map[string]any)
	if !ok || originalStart["dateTime"] != "2026-01-02T10:00:00Z" {
		t.Fatalf("cancelled instance originalStartTime=%#v", instance["originalStartTime"])
	}
}

func TestCalendarEventsListCall_EventTypesFilter(t *testing.T) {
	cases := []struct {
		name       string
		eventTypes []string
		want       []string
	}{
		// No filter: the eventTypes query param must be absent, preserving the
		// API default of returning all event types (non-breaking).
		{"unset omits the filter", nil, nil},
		// Filter: the requested types are sent as repeated eventTypes params.
		{"sends requested types", []string{eventTypeBirthday, eventTypeDefault}, []string{eventTypeBirthday, eventTypeDefault}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			var gotTypes []string
			srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
				gotTypes = r.URL.Query()["eventTypes"]
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

			if _, err := calendarEventsListCall(context.Background(), svc, "primary", "2026-01-01T00:00:00Z", "2026-01-02T00:00:00Z", 10, "", "", "", "", tc.eventTypes, "", false).Do(); err != nil {
				t.Fatalf("Do: %v", err)
			}
			if !slices.Equal(gotTypes, tc.want) {
				t.Fatalf("eventTypes query = %v, want %v", gotTypes, tc.want)
			}
		})
	}
}
