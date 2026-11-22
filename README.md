# Recurring Task Management System

An automated task management system built using **Google Sheets and Google Apps Script** to manage recurring tasks, track employee submissions, monitor overdue work, and send automated alerts.

## Features

- **Recurring Task Management:** Automatically generates the next task cycle after a task is completed.
- **Task Tracking:** Tracks task status, due dates, completion proof, and submission details.
- **Manager Dashboard:** Provides an overview of active, pending, overdue, and completed tasks.
- **Automated Manager Alerts:** Sends email notifications when tasks remain incomplete after their due time.
- **Leadership Escalation:** Escalates tasks that remain incomplete beyond a defined time threshold.
- **Alert Acknowledgement:** Allows managers to acknowledge alerts and track their resolution status.
- **Automated Scheduling:** Uses time-driven triggers to check overdue tasks and escalations.

## Technology Stack

- Google Sheets
- Google Apps Script
- JavaScript
- Gmail notifications
- Time-driven triggers

## How It Works

1. Tasks are configured in the task master sheet.
2. The system generates recurring task cycles.
3. Employees submit completion details and supporting proof.
4. The system updates task statuses and creates the next cycle.
5. Managers receive alerts for overdue tasks.
6. Unresolved tasks are escalated to leadership.
7. Alerts and acknowledgements are recorded for tracking.

## Project Structure

```text
Recurring-Task-Management-System/
├── Code.gs
├── README.md
├── test-cases.md
└── screenshots/
```

## Setup

1. Create a Google Spreadsheet.
2. Open **Extensions → Apps Script**.
3. Add the contents of `Code.gs`.
4. Configure the spreadsheet tabs and columns expected by the script.
5. Set the manager and leadership email addresses in the script.
6. Run the initialization function and authorize the required permissions.
7. Enable the required time-driven triggers.

## Testing

The system can be tested using sample employee names, placeholder email addresses, and dummy task records.

## Note

This is a portfolio project demonstrating workflow automation, recurring task management, monitoring, and escalation using Google Apps Script and Google Sheets.