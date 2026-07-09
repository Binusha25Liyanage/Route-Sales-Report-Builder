---
name: Cherry Alloy Enterprise
colors:
  surface: '#fcf9f4'
  surface-dim: '#dcdad5'
  surface-bright: '#fcf9f4'
  surface-container-lowest: '#ffffff'
  surface-container-low: '#f6f3ee'
  surface-container: '#f0ede9'
  surface-container-high: '#ebe8e3'
  surface-container-highest: '#e5e2dd'
  on-surface: '#1c1c19'
  on-surface-variant: '#544242'
  inverse-surface: '#31302d'
  inverse-on-surface: '#f3f0eb'
  outline: '#877272'
  outline-variant: '#dac1c0'
  surface-tint: '#974449'
  primary: '#5d181e'
  on-primary: '#ffffff'
  primary-container: '#7a2e33'
  on-primary-container: '#ff999c'
  inverse-primary: '#ffb3b4'
  secondary: '#615e59'
  on-secondary: '#ffffff'
  secondary-container: '#e5dfd9'
  on-secondary-container: '#66625e'
  tertiary: '#35312a'
  on-tertiary: '#ffffff'
  tertiary-container: '#4c473f'
  on-tertiary-container: '#bcb5ab'
  error: '#ba1a1a'
  on-error: '#ffffff'
  error-container: '#ffdad6'
  on-error-container: '#93000a'
  primary-fixed: '#ffdad9'
  primary-fixed-dim: '#ffb3b4'
  on-primary-fixed: '#3f020b'
  on-primary-fixed-variant: '#792d33'
  secondary-fixed: '#e7e1dc'
  secondary-fixed-dim: '#cbc6c0'
  on-secondary-fixed: '#1d1b18'
  on-secondary-fixed-variant: '#494642'
  tertiary-fixed: '#eae1d7'
  tertiary-fixed-dim: '#cdc5bc'
  on-tertiary-fixed: '#1f1b15'
  on-tertiary-fixed-variant: '#4b463f'
  background: '#fcf9f4'
  on-background: '#1c1c19'
  surface-variant: '#e5e2dd'
typography:
  headline-lg:
    fontFamily: Manrope
    fontSize: 32px
    fontWeight: '700'
    lineHeight: 40px
    letterSpacing: -0.02em
  headline-md:
    fontFamily: Manrope
    fontSize: 24px
    fontWeight: '600'
    lineHeight: 32px
    letterSpacing: -0.01em
  headline-sm:
    fontFamily: Manrope
    fontSize: 18px
    fontWeight: '600'
    lineHeight: 24px
  body-lg:
    fontFamily: Work Sans
    fontSize: 16px
    fontWeight: '400'
    lineHeight: 24px
  body-md:
    fontFamily: Work Sans
    fontSize: 14px
    fontWeight: '400'
    lineHeight: 20px
  body-sm:
    fontFamily: Work Sans
    fontSize: 12px
    fontWeight: '400'
    lineHeight: 16px
  label-caps:
    fontFamily: IBM Plex Sans
    fontSize: 11px
    fontWeight: '600'
    lineHeight: 16px
    letterSpacing: 0.1em
  table-numeric:
    fontFamily: Work Sans
    fontSize: 14px
    fontWeight: '500'
    lineHeight: 20px
rounded:
  sm: 0.25rem
  DEFAULT: 0.5rem
  md: 0.75rem
  lg: 1rem
  xl: 1.5rem
  full: 9999px
spacing:
  base: 4px
  xs: 4px
  sm: 8px
  md: 16px
  lg: 24px
  xl: 32px
  container-margin: 40px
  gutter: 20px
---

## Brand & Style

This design system is engineered for high-performance internal business environments where data density and clarity are paramount. The aesthetic merges **Corporate Modernism** with a **Tactile** edge, utilizing a sophisticated "Cherry Alloy" palette to evoke a sense of heritage, stability, and industrial precision.

The target audience consists of analysts and decision-makers who require rapid information retrieval without cognitive fatigue. The UI avoids the stark coldness of typical enterprise software by utilizing a warm off-white base, creating an environment that feels premium yet utilitarian. Visual hierarchy is established through meticulous typography and subtle structural depth rather than aggressive color blocks.

## Colors

The palette is anchored by **Cherry Alloy (#7A2E33)**, a deep, metallic wine-red used sparingly for primary actions, branding, and critical data points. The background uses **Parchment White (#FAF7F2)** to reduce eye strain during long working sessions.

- **Primary:** Cherry Alloy (#7A2E33) — Reserved for CTA buttons, active states, and key brand identifiers.
- **Secondary/Neutral High:** Stone Charcoal (#4A4743) — Used for primary headings and prominent icons.
- **Neutral Low:** Warm Slate (#D9D1C7) — Used for borders and disabled states.
- **Background:** Off-White (#FAF7F2) — The canvas for all application surfaces.
- **Functional:** Success (Forest), Warning (Amber), and Error (Crimson) should be desaturated to maintain the professional, understated tone.

## Typography

The typographic system prioritizes legibility and information architecture. 

- **Headlines:** Use **Manrope** for its modern, balanced proportions. It provides a confident, clean look for page titles and section headers.
- **Body & Data:** **Work Sans** is utilized for its exceptional performance in data-heavy environments. All numerical data must use **tabular figures** (monospaced numbers) to ensure columns align perfectly in financial or analytical views.
- **Section Labels:** Small uppercase labels using **IBM Plex Sans** with increased tracking (10%) are used to categorize content blocks without adding visual weight.
- **Scale:** On mobile devices, `headline-lg` scales down to 24px (`headline-lg-mobile`) to prevent awkward line breaks in narrow viewports.

## Layout & Spacing

This design system employs a **Fixed Grid** philosophy for desktop layouts to maintain a disciplined, "dashboard" feel, transitioning to a fluid model for smaller breakpoints.

- **Desktop (1440px+):** 12-column grid with a max-width of 1280px. Gutters are fixed at 20px.
- **Tablet (768px - 1439px):** 8-column fluid grid with 16px margins.
- **Mobile (Up to 767px):** 4-column fluid grid with 16px margins.

We utilize a **4px baseline grid**. Components like cards and tables should favor "generous" internal padding (`lg` or 24px) to balance the high density of the data contained within. Use white space as a structural tool to separate unrelated data clusters.

## Elevation & Depth

Hierarchy is established through **Tonal Layering** and **Ambient Shadows**. Instead of heavy dropshadows, we use a multi-layered approach to make elements feel integrated into the surface.

- **Surface Level (Level 0):** The main background (#FAF7F2).
- **Card Level (Level 1):** White (#FFFFFF) surfaces with a 1px border in Stone (#D9D1C7) and a very soft, diffused shadow (0px 4px 12px rgba(74, 71, 67, 0.05)).
- **Popovers/Modals (Level 2):** Same as Level 1 but with a more pronounced shadow (0px 12px 24px rgba(74, 71, 67, 0.12)) to indicate temporary interaction.

Avoid pure black shadows; always tint shadows with the secondary stone color to maintain the warmth of the off-white canvas.

## Shapes

The design system utilizes **Rounded (0.5rem)** corners as the standard for all primary containers, providing a professional yet approachable feel.

- **Standard Cards:** 0.5rem (8px) corner radius.
- **Buttons & Inputs:** 0.5rem (8px) for consistency with cards.
- **Small Components (Chips/Tags):** 0.25rem (4px) to maintain sharpness at smaller scales.
- **Status Indicators:** Perfect circles for notifications or status dots.

## Components

### Buttons
- **Primary:** Solid Cherry Alloy (#7A2E33) with white text. High-contrast, bold presence.
- **Secondary:** Transparent with a Stone Charcoal (#4A4743) border and text.
- **Tertiary:** Ghost style; no border, Stone Charcoal text, subtle gray hover state.

### Input Fields
- Background: White (#FFFFFF).
- Border: 1px Solid Stone (#D9D1C7). 
- Active State: 1px Solid Cherry Alloy with a 2px soft glow in the same color.
- Labels: `label-caps` typography, positioned above the field.

### Cards
- Always use **rounded-lg** (1rem/16px) for main dashboard containers.
- Padding: 24px (`lg`).
- Header: Include a subtle 1px bottom border to separate the title area from the content.

### Data Tables
- Row Height: 48px for standard density; 40px for high density.
- Header: `label-caps` style, muted stone gray background (#F4F1ED).
- Typography: Use `table-numeric` for all value-based columns.

### Icons
- Style: Lucide-inspired outline icons.
- Stroke Weight: 1.5px.
- Size: 20px for primary navigation; 16px for inline actions and labels.