# Prototype Instructions

Run the local server yourself and open the preview in the browser available to this environment. Do not give the user server-start instructions when you can run it.

Before making substantial visual changes, use the Product Design plugin's `get-context` skill when the visual source is unclear or no longer matches the current goal. When the user gives durable prototype-specific design feedback, preferences, or decisions, record them in `AGENTS.md`.

When implementing from a selected generated mock, treat that image as the source of truth for layout, component anatomy, density, spacing, color, typography, visible content, and hierarchy.

## Product decisions

- Build an independent Korean school exam-management app grounded in the captured `hirame-ki.github.io/exam` desktop flow.
- Use the source navy/gold visual language, Pretendard UI text, and a Korean serif display face.
- Fix the source's mobile horizontal overflow with a single-column responsive layout and horizontally scrollable navigation.
- The production backend is a school-owned Google Sheet with bound Apps Script. The frontend must never contain real staff login names, PINs, administrator passwords, student names, or backend secrets.
- Teacher access uses an administrator-managed login name plus a six-digit personal PIN. Headquarters, seating, setup, destructive actions, and exports require an administrator session.
- The login page exposes no staff-name directory. A separate administrator entry allows initial access-user setup and recovery when no users are active.
- Use `한양대학교사범대학부속고등학교 통합 시스템` as the product name across the login page, header, browser metadata, demo data, and public configuration.
- Do not show demo labels, demo links, infrastructure copy, or a product footer in the application UI.
- `?demo=1` must use synthetic data only and must never send writes to the production endpoint.
- Display exam dates with the full Korean weekday throughout the UI, and show the weekday immediately beside the administrator's date input.
- Keep the connection count and five-minute idle-session notice together in the upper-right header; shorten only the visible idle label to `5분` at 560px and below.
- Use `한양대학교사범대학부속고등학교` as the synthetic demo school name and omit decorative English eyebrow text from every screen.
- Timetable subjects use a per-grade Excel-backed recommendation catalog, but administrators may still type an unlisted subject directly. The `전체` class toggle expands to real active class IDs when saved.
- The parent notice combines grades 1–3 on one A4 landscape page using all active exam dates, with editable greeting, notes, issue date, and issuer.
