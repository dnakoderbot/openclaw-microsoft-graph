import { graphFetch } from "./graph-client.js";
import type { ResolvedOutlookAccount } from "./types.js";

export type GraphCalendarEvent = {
  id?: string;
  subject?: string;
  webLink?: string;
  start?: {
    dateTime?: string;
    timeZone?: string;
  };
  end?: {
    dateTime?: string;
    timeZone?: string;
  };
  location?: {
    displayName?: string;
  };
  attendees?: Array<{
    emailAddress?: {
      name?: string;
      address?: string;
    };
    type?: string;
  }>;
};

export async function listUpcomingEvents(params: {
  account: ResolvedOutlookAccount;
  top?: number;
  startDateTime?: string;
  endDateTime?: string;
}): Promise<GraphCalendarEvent[]> {
  const startDateTime = params.startDateTime ?? new Date().toISOString();
  const endDateTime =
    params.endDateTime ?? new Date(Date.now() + 7 * 24 * 60 * 60_000).toISOString();
  const top = params.top ?? 10;
  const response = await graphFetch(
    params.account,
    `/me/calendarView?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}&$top=${top}&$orderby=start/dateTime&$select=id,subject,webLink,start,end,location,attendees`,
  );
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph calendarView failed: HTTP ${response.status} ${body}`);
  }
  const payload = (await response.json()) as { value?: GraphCalendarEvent[] };
  return payload.value ?? [];
}

export async function createCalendarEvent(params: {
  account: ResolvedOutlookAccount;
  subject: string;
  bodyText?: string;
  startDateTime: string;
  endDateTime: string;
  timeZone?: string;
  attendeeEmails?: string[];
  locationDisplayName?: string;
}): Promise<GraphCalendarEvent> {
  const timeZone = params.timeZone ?? "UTC";
  const response = await graphFetch(params.account, "/me/events", {
    method: "POST",
    body: JSON.stringify({
      subject: params.subject,
      body: params.bodyText
        ? {
            contentType: "Text",
            content: params.bodyText,
          }
        : undefined,
      start: {
        dateTime: params.startDateTime,
        timeZone,
      },
      end: {
        dateTime: params.endDateTime,
        timeZone,
      },
      location: params.locationDisplayName
        ? {
            displayName: params.locationDisplayName,
          }
        : undefined,
      attendees: (params.attendeeEmails ?? []).map((address) => ({
        emailAddress: { address },
        type: "required",
      })),
    }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph create event failed: HTTP ${response.status} ${body}`);
  }
  return (await response.json()) as GraphCalendarEvent;
}
