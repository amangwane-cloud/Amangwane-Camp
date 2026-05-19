# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Static HTML website for **Amangwane Tented Camp** — a boutique accommodation at Kosi Bay Mouth, iSimangaliso Wetland Park, South Africa. No build system, no framework, no package manager. All files are served directly as-is.

**Live site:** https://amangwanecamp.co.za

## Development

No build step required. Open any `.html` file directly in a browser, or serve the directory with any static file server:

```powershell
# Python (if installed)
python -m http.server 8000

# Node (if installed)
npx serve .
```

## Site Structure

| File | Purpose |
|------|---------|
| `index.html` | Homepage — hero, about, accommodation preview, testimonials, activities, contact form |
| `accommodation.html` | Full accommodation listings with amenities |
| `activities.html` | Activities (boat cruise, snorkeling, turtle tours, kayaking, fishing, birding) |
| `blog.html` | Travel guides and Kosi Bay articles |
| `gallery.html` | Photo gallery |

## Architecture

Every page is a single self-contained HTML file: styles in `<style>` blocks, scripts in `<script>` blocks at the bottom. No external local CSS/JS files.

**Shared patterns across all pages:**
- Fixed top nav with `openMnav()` / `closeMnav()` for the mobile burger menu
- Hero section with background image and `.hero-overlay`
- Scroll-triggered animations: elements get a `reveal` class, JavaScript adds `visible` on scroll
- Analytics events tracked on key CTA clicks (Book Now, WhatsApp, email, phone)

## Design Tokens (CSS Variables)

```css
--sun:    #F5C518   /* yellow — primary CTAs */
--ocean:  #1A8CB5   /* blue — secondary elements */
--forest: #2A6E2A   /* green — tertiary */
--charcoal:#1C1C2A  /* dark — primary text/backgrounds */
--sand:   #E8D5A0   /* warm neutral — accents */
--display:'Fraunces'  /* serif — headings */
--body:   'DM Sans'   /* sans-serif — body copy */
```

## Third-Party Integrations

| Service | ID / Account | Purpose |
|---------|-------------|---------|
| Google Analytics 4 | `G-526493660` | Traffic analytics |
| Google Tag Manager | `GTM-MMQ98GGR` | Event management |
| Facebook Pixel | `691106136867109` | Retargeting |
| InnStyle | `amangwanetentedcamp` | Booking calendar (`https://developer.innstyle.co.uk/calendar.min.js`) |

## Key Implementation Notes

- All IDs and tracking codes are hardcoded directly in HTML — no environment variables.
- `accommodation.html` and `gallery.html` are very large files (7 MB and 16 MB respectively) due to base64-encoded or inline images. Be careful with full-file reads; target specific sections when editing.
- Contact form uses a `submitContactForm()` JavaScript function that posts to an external service and shows an inline success message.
- SEO: each page has canonical `<link>` tags and structured meta tags — keep these consistent when adding or renaming pages.
