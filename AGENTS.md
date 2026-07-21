# Prototype Instructions

Run the local server yourself and open the preview in the browser available to this environment. Do not give the user server-start instructions when you can run it.

Before making substantial visual changes, use the Product Design plugin's `get-context` skill when the visual source is unclear or no longer matches the current goal. When the user gives durable prototype-specific design feedback, preferences, or decisions, record them in `AGENTS.md`.

When implementing from a selected generated mock, treat that image as the source of truth for layout, component anatomy, density, spacing, color, typography, visible content, and hierarchy.

## Product decisions

- Build an independent Korean school exam-management app grounded in the captured `hirame-ki.github.io/exam` desktop flow.
- Use the source navy/gold visual language, Pretendard UI text, and a Korean serif display face.
- Fix the source's mobile horizontal overflow with a single-column responsive layout and horizontally scrollable navigation.
- The production backend is a school-owned Google Sheet with bound Apps Script. The frontend must never contain the school code, administrator password, student names, or backend secrets.
- Teacher access uses a school-code session. Headquarters, seating, setup, destructive actions, and exports require an administrator session.
- `?demo=1` must use synthetic data only and must never send writes to the production endpoint.
- Display exam dates with the full Korean weekday throughout the UI, and show the weekday immediately beside the administrator's date input.
- Keep the connection count and five-minute idle-session notice together in the upper-right header; shorten only the visible idle label to `5분` at 560px and below.
- Use `한양대학교사범대학부속고등학교` as the synthetic demo school name and omit decorative English eyebrow text from every screen.
- Timetable subjects use a per-grade Excel-backed recommendation catalog, but administrators may still type an unlisted subject directly. The `전체` class toggle expands to real active class IDs when saved.
- The parent notice combines grades 1–3 on one A4 landscape page using all active exam dates, with editable greeting, notes, issue date, and issuer.
