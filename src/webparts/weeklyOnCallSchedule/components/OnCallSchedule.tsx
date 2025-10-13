import * as React from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { startOfWeek, endOfWeek, addDays, format } from 'date-fns';

export interface IOnCallScheduleProps {
  context: any;
}

interface ICalendarEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone?: string };
  end: { dateTime: string; timeZone?: string };
  organizer?: { emailAddress?: { name?: string; address?: string } };
  location?: { displayName?: string };
  attendees?: {
    emailAddress?: { name?: string; address?: string };
    type?: string;
    status?: { response?: string; time?: string };
  }[];
}

const OnCallSchedule: React.FC<IOnCallScheduleProps> = ({ context }) => {
  const [events, setEvents] = React.useState<ICalendarEvent[]>([]);
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    const fetchEvents = async (): Promise<void> => {
      try {
        const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');

        const start = startOfWeek(new Date(), { weekStartsOn: 0 });
        const end = endOfWeek(new Date(), { weekStartsOn: 0 });

        const res = await client
          .api(`/users/it-oncall@cvch.org/calendarView`)
          .query({
            startDateTime: start.toISOString(),
            endDateTime: end.toISOString(),
          })
          .select('subject,start,end,attendees')
          .orderby('start/dateTime')
          .get();

        console.log('Graph Results:', res.value);
        setEvents(res.value || []);
      } catch (err) {
        console.error('Graph error:', err);
      } finally {
        setLoading(false);
      }
    };

    fetchEvents();
  }, [context]);

  if (loading) return <div>Loading on-call schedule…</div>;

  const today = new Date();
  const weekStart = startOfWeek(today, { weekStartsOn: 0 });
  const weekEnd = endOfWeek(today, { weekStartsOn: 0 });

  const days = Array.from({ length: 7 }).map((_, i) => {
    const date = addDays(weekStart, i);

    const event = events.find((ev) => {
      const startDate = new Date(ev.start.dateTime + 'Z');
      const endDate = new Date(ev.end.dateTime + 'Z');
      return date >= startDate && date < endDate;
    });

    return {
      day: format(date, 'EEE'),
      date: format(date, 'MMM d'),
      name: event?.attendees?.[0]?.emailAddress?.name || '—',
      highlight: date.toDateString() === today.toDateString(),
    };
  });

  return (
    <div
      style={{
        backgroundColor: '#5B1647',
        borderRadius: '1rem',
        padding: '2rem 0',
        color: 'white',
        width: '100%',
      }}
    >
      <div
        style={{
          maxWidth: '1200px',
          margin: '0 auto',
          padding: '0 1.5rem',
        }}
      >
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            flexWrap: 'wrap',
            gap: '1rem',
          }}
        >
          <h3 style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
            📅 Weekly On-Call Schedule
          </h3>
          <p style={{ color: '#e0cbe0', margin: 0 }}>
            Week of {format(weekStart, 'MMM d')} – {format(weekEnd, 'MMM d')}
          </p>
        </div>

        <div
          style={{
            display: 'flex',
            gap: '1rem',
            marginTop: '1.5rem',
            flexWrap: 'nowrap',
            overflowX: 'auto',
            justifyContent: 'center',
            scrollSnapType: 'x mandatory',
            paddingBottom: '0.5rem',
          }}
        >
          {days.map((d: any) => (
            <div
              key={d.day}
              style={{
                backgroundColor: d.highlight ? '#c0187a' : '#6A1B5A',
                borderRadius: '0.75rem',
                padding: '1rem',
                flex: '0 0 120px',
                textAlign: 'center',
                boxShadow: d.highlight ? '0 0 10px rgba(255,255,255,0.3)' : 'none',
                scrollSnapAlign: 'center',
              }}
            >
              <div style={{ fontWeight: 700 }}>{d.day}</div>
              <div
                style={{
                  fontSize: '0.9rem',
                  color: '#e0cbe0',
                  borderBottom: '1px solid #e0cbe0',
                  margin: '0.25rem 0 0.5rem',
                }}
              >
                {d.date}
              </div>
              <div style={{ fontWeight: 600 }}>{d.name}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default OnCallSchedule;
