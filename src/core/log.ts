/**
 * The only logging in the add-in.
 *
 * Privacy rule: never log file contents, file names, subjects, recipients,
 * message bodies or item IDs. The type only accepts a fixed event code plus
 * numbers and booleans, so a string cannot be passed in by accident.
 */
export type LogCode =
  | "start"
  | "unsupported-host"
  | "busy"
  | "no-item"
  | "plan"
  | "read-done"
  | "confirm-shown"
  | "confirm-unavailable"
  | "cancelled"
  | "selection-changed"
  | "opened"
  | "open-failed"
  | "error";

export type LogData = Readonly<Record<string, number | boolean>>;

export function logEvent(code: LogCode, data: LogData = {}): void {
  console.info(`[reply-all-with-attachments] ${code}`, JSON.stringify(data));
}
