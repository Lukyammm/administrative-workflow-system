# Locker Custody Workflow

Repository: `locker-custody-workflow`

## Overview

Administrative web app for controlled locker usage, custody terms, user sessions, notifications, and audit trails.

## Main Capabilities

- Cosign-style sidebar with dashboard, lockers, terms, notifications, and audit views.
- Locker state monitoring for in-use, available, expiring, contingency, and overdue records.
- Draft persistence for custody terms and locker movements.
- Header preferences for profile, unit, theme, avatar, and session state.

## Operating Flow

1. The user signs in and selects profile and unit context.
2. The dashboard shows locker availability and overdue items.
3. The term workflow collects required fields, saves drafts, and registers the movement.
4. Notifications and audit entries keep the operation traceable.

## Visual System Guide

> The screens below are documentation mockups based on the components, labels, colors, and workflows found in this repository. All displayed data is fictitious and does not represent real patients, staff members, or institutions.

### Cosign - locker dashboard
![Cosign - locker dashboard](./docs/screenshots/dashboard.svg)

### Cosign - terms and release workflow
![Cosign - terms and release workflow](./docs/screenshots/workflow-board.svg)

### Cosign - audit and session state
![Cosign - audit and session state](./docs/screenshots/audit-log.svg)

## Data Privacy

The repository documentation and guide images use fictitious sample data only.

## Technologies

- JavaScript
- HTML/CSS
- Google Apps Script
- Google Sheets

## Status

Completed
