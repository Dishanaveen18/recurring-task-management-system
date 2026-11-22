# Test Cases – Recurring Task Management System

## 1. Task Initialization
- **Test:** Run the initialization function.
- **Expected Result:** Initial task cycles should be generated from the task master sheet.
- **Status:** Passed

## 2. Employee Task Completion
- **Test:** Submit a task with the required completion proof.
- **Expected Result:** Task status should change to Completed.
- **Status:** Passed

## 3. Recurring Task Generation
- **Test:** Complete a recurring task.
- **Expected Result:** The system should automatically generate the next task cycle based on its frequency.
- **Status:** Passed

## 4. Manager Overdue Alert
- **Test:** Leave a task incomplete beyond the configured overdue threshold.
- **Expected Result:** The manager should receive an email notification, and the alert should be recorded in the Alert Log.
- **Status:** Passed

## 5. Alert Acknowledgement
- **Test:** Acknowledge a manager alert.
- **Expected Result:** The acknowledgement time, user, and resolution status should be recorded.
- **Status:** Passed

## 6. Leadership Escalation
- **Test:** Leave a task unresolved beyond the leadership escalation threshold.
- **Expected Result:** Leadership should receive an escalation email, and the event should be recorded in the Alert Log.
- **Status:** Passed

## 7. Automated Triggers
- **Test:** Configure time-driven triggers.
- **Expected Result:** The system should automatically check overdue tasks and leadership escalations at the configured intervals.
- **Status:** Passed

*Note: These test cases reflect the features tested during development. Use anonymized sample data when demonstrating the project.*