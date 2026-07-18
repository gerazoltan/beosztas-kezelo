import type { CalendarEvent } from '../domain/types';

export function calendarEventDescription(item: CalendarEvent): string {
  const lines = [`Szolgálati jelleg: ${item.serviceCategory}`];
  if (item.inference) {
    lines.push(
      'Felismerés: a szolgálattípus a napi szolgálati összeállításból lett következtetve.',
      `Technikai magyarázat: ${item.inference.explanation}`,
    );
  }
  return lines.join('\n');
}
